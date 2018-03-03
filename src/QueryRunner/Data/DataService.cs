using QueryRunner.Data.Entities;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using QueryRunner.Utilities;
using System.Data;
using Microsoft.Office.Interop.Access.Dao;

namespace QueryRunner.Data
{
    public class DataService
    {
        private string _connectionString;
        private string _databasePath = string.Empty;
        private OleDbConnection _connection;

        private DataService(string databasePath, ref List<string> messages)
        {
            _databasePath = databasePath;
            SetConnection(ref messages);
        }

        public static DataService CreateDataService(string databasePath, out List<string> messages)
        {
            messages = new List<string>();
            if (VerifyFile(databasePath))
            {
                return new DataService(databasePath, ref messages);
            }
            messages.Add("Specified database file could not be found or extension is not supported.");
            return null;
        }

        private bool SetConnection(ref List<string> messages)
        {
            bool returnValue = false;
            // Sets the database connection based on Access 2013/2016.
            _connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+ _databasePath + ";Persist Security Info=False;";

            if (_connection != null)
            {
                _connection.Close();
                _connection.Dispose();
            }

            try
            {
                _connection = new OleDbConnection(_connectionString);
                messages.Add("Database connection set successfully.");
                returnValue = true;
            }
            catch (Exception ex)
            {
                messages.Add(ex.Message);
            }

            return returnValue;
        }

        public bool ResetConnection(string databasePath, out List<string> messages)
        {
            messages = new List<string>();
            if (_databasePath == databasePath)
            {
                messages.Add("Connection based on specified database is already set.");
                return true;
            }

            if (VerifyFile(databasePath))
            {
                _databasePath = databasePath;
                return SetConnection(ref messages);
            }

            messages.Add("Specified database file could not be found or extension is not supported.");
            messages.Add("Connection was not reset.");

            return false;

        }

        private static bool VerifyFile(string filePath, string fileExtension = ".accdb")
        {
            if (string.IsNullOrWhiteSpace(filePath)) return false;

            if (System.IO.File.Exists(filePath))
            {
                fileExtension = fileExtension.StartsWith(".") ? fileExtension : "." + fileExtension;
                string extension = System.IO.Path.GetExtension(filePath);
                if (extension == fileExtension)
                {
                    return true;
                }
            }
            return false;
        }

        public string ReleaseComObject(object comObject)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                return string.Format("Error encountered releasing COM object: {0}", ex.Message);
            }
            finally
            {
                comObject = null;
            }

            return "COM object released.";
        }

        public List<Query> GetQueryList(DateTime defaultStartDate, DateTime defaultEndDate, out List<string> messages)
        {
            messages = new List<string>();

            DataTable viewsTable = null;
            DataTable proceduresTable = null;

            List<Query> queries = new List<Query>();
            List<Query> procedures = new List<Query>();

            try
            {
                _connection.Open();
                viewsTable = _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Views, null);

                foreach (DataRow row in viewsTable.Rows)
                {
                    Query query = new Query
                    {
                        QueryName = row["TABLE_NAME"].ToString(),
                        QueryDefinition = row["VIEW_DEFINITION"].ToString(),
                        QueryCommandType = CommandType.TableDirect
                    };
                    queries.Add(query);
                }

                proceduresTable = _connection.GetOleDbSchemaTable(OleDbSchemaGuid.Procedures, null);
                foreach (DataRow row in proceduresTable.Rows)
                {
                    Query query = new Query
                    {
                        QueryName = row["PROCEDURE_NAME"].ToString(),
                        QueryDefinition = row["PROCEDURE_DEFINITION"].ToString(),
                        QueryCommandType = CommandType.StoredProcedure
                    };
                    procedures.Add(query);
                }
            }
            catch (Exception ex)
            {
                messages.Add(ex.Message);
            }
            finally
            {
                _connection.Close();
                if (viewsTable != null)
                {
                    viewsTable.Dispose();
                }
                if (proceduresTable != null)
                {
                    proceduresTable.Dispose();
                }
            }

            if (procedures.Count > 0)
            {
                DBEngine engine = new DBEngine();
                Database database = null;
                Parameters parameters = null;

                database = engine.OpenDatabase(_databasePath);
                QueryDefs defs = database.QueryDefs;
                QueryDef def = null;

                foreach (Query procedure in procedures)
                {
                    def = defs[procedure.QueryName];

                    parameters = def.Parameters;
                    int count = 0;

                    try
                    {
                        count = parameters.Count;
                    }
                    catch (Exception ex)
                    {
                        messages.Add(string.Format("Error parsing '{0}': {1} The '{0}' procedure is not valid outside of the Access environment.", procedure.QueryName, ex.Message));
                        procedure.Valid = false;
                    }

                    for (int i = 0; i < count; i++)
                    {
                        QueryParameter qp = new QueryParameter
                        {
                            ParameterName = parameters[i].Name,
                            Type = TypeMapper.MapDaoToOleDbType((DataTypeEnum)parameters[i].Type)
                        };

                        if (qp.ParameterName == "[Start_Date]")
                        {
                            qp.Value = defaultStartDate.ToShortDateString();
                        }
                        if (qp.ParameterName == "[End_Date]")
                        {
                            qp.Value = defaultEndDate.ToShortDateString();
                        }

                        procedure.QueryParameters.Entities.Add(qp);
                    }

                    def.Close();
                }

                if (database != null)
                {
                    database.Close();
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                ReleaseComObject(parameters);
                ReleaseComObject(defs);
                ReleaseComObject(def);

                if (database != null)
                {
                    ReleaseComObject(database);
                }

                ReleaseComObject(engine);

                queries.AddRange(procedures);
            }

            return queries;
        }

        public DataTable GetResultsTable(Query query, out List<string> messages)
        {
            messages = new List<string>();
            DataTable readerTable = null;

            _connection.Open();

            using (OleDbCommand command = new OleDbCommand())
            {
                command.Connection = _connection;

                command.CommandType = query.QueryCommandType;

                if (query.QueryCommandType == CommandType.Text)
                {
                    command.CommandText = query.QueryDefinition;
                }

                if (query.QueryCommandType == CommandType.TableDirect)
                {
                    command.CommandText = "[" + query.QueryName + "]";
                }

                if (query.QueryCommandType == CommandType.StoredProcedure)
                {
                    command.CommandText = "[" + query.QueryName + "]";

                    if (query.QueryParameters.Entities.Count > 0)
                    {
                        foreach (QueryParameter parameter in query.QueryParameters.Entities)
                        {
                            OleDbParameter p = new OleDbParameter(parameter.ParameterName, parameter.Type);

                            if (parameter.Value != null)
                            {

                                Type type = TypeMapper.MapOleDbTypeToCLR(parameter.Type);
                                try
                                {
                                    p.Value = Convert.ChangeType(parameter.Value, type);
                                }
                                catch (Exception ex)
                                {
                                    messages.Add(string.Format("Invalid data specified for parameter {0}. No value passed to query. ({1})", p.ParameterName, ex.Message));
                                    p.Value = DBNull.Value;
                                }

                            }
                            else
                            {
                                p.Value = DBNull.Value;
                            }
                            command.Parameters.Add(p);
                        }
                    }
                }

                try
                {
                    using (OleDbDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection))
                    {
                        if (reader.HasRows)
                        {
                            readerTable = new DataTable(query.QueryName);

                            for (int count = 0; count < reader.FieldCount; count++)
                            {
                                DataColumn column = new DataColumn(reader.GetName(count), reader.GetFieldType(count));
                                readerTable.Columns.Add(column);
                            }

                            while (reader.Read())
                            {
                                DataRow dr = readerTable.NewRow();
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    dr[i] = reader.GetValue(reader.GetOrdinal(reader.GetName(i)));
                                }
                                readerTable.Rows.Add(dr);
                            }
                        }

                        reader.Close();
                    }
                }
                catch (Exception ex)
                {
                    messages.Add(ex.Message);
                }
                finally
                {
                    _connection.Close();
                }

                command.Parameters.Clear();
                messages.Add(string.Format("The '{0}' query was processed.", query.QueryName));
            }

            return readerTable;
        }
    }
}
