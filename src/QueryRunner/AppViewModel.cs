using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using QueryRunner.Data;
using QueryRunner.Data.Entities;
using QueryRunner.Utilities;

namespace QueryRunner
{
    public class AppViewModel : INotifyPropertyChanged
    {
        private string _databasePath = string.Empty;
        private string _textFileDirectory = string.Empty;
        private string _excelFilePath = string.Empty;
        private bool _textOutput = true;

        private bool _rememberDatabase = true;
        private DataService _dataService;

        private ObservableCollection<string> _messages;

        private ObservableCollection<Query> _queries;
        private List<Query> _selectedQueries;

        // Date range defaults to current week.
        private DateTime _startDate = DateTimeExtensions.StartOfWeek(DateTime.Today, DayOfWeek.Monday);
        private DateTime _endDate = DateTimeExtensions.EndOfWeek(DateTime.Today, DayOfWeek.Monday);
        private string _processTime = string.Empty;
        private string _statusMessage = string.Empty;
        private bool _idle = true;

        private RelayCommand _browseForDatabaseCommand;
        private RelayCommand _browseForFileCommand;
        private RelayCommand _browseForDirectoryCommand;
        private RelayCommand _setDateRangeCommand;
        private RelayCommand _runQueriesCommand;
        private RelayCommand _openDirectoryCommand;
        private RelayCommand _clearMessagesCommand;

        public AppViewModel()
        {
            _messages = new ObservableCollection<string>();
            DatabasePath = Properties.Settings.Default.DatabasePath;

            TextFileDirectory = Properties.Settings.Default.TextFileDirectory;
            ExcelFilePath = Properties.Settings.Default.ExcelFilePath;
            RememberDatabase = Properties.Settings.Default.RememberDatabase;
        }

        #region INotifyPropertyChanged Implementation ---------------------------------------------

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion // INotifyPropertyChanged Implementation

        #region Properties ------------------------------------------------------------------------
        public ObservableCollection<string> Messages
        {
            get
            {
                return _messages;
            }
            private set { }
        }

        public bool TextOutput
        {
            get
            {
                return _textOutput;
            }
            set
            {
                _textOutput = value;
                OnPropertyChanged(propertyName: nameof(TextOutput));
            }
        }

        public string DatabasePath
        {
            get
            {
                return _databasePath;
            }
            set
            {
                if (_databasePath != value)
                {
                    string _value = value;
                    bool connectionSet = false;

                    List<string> messages = new List<string>();
                    if (_dataService == null)
                    {
                        _dataService = DataService.CreateDataService(_value, out messages);
                        connectionSet = (_dataService != null);        
                    }
                    else
                    {
                        connectionSet = _dataService.ResetConnection(_value, out messages);
                    }
                    foreach (string message in messages)
                    {
                        Messages.Add(message);
                    }

                    if (connectionSet)
                    {                        
                        _databasePath = _value;
                        if (RememberDatabase == true)
                        {
                            Properties.Settings.Default.DatabasePath = _databasePath;
                            Properties.Settings.Default.Save();
                        }
                        OnPropertyChanged(propertyName: nameof(DatabasePath));
                        _queries = null;
                        OnPropertyChanged(propertyName: nameof(Queries));
                    }
                }
            }
        }

        public bool RememberDatabase
        {
            get
            {
                return _rememberDatabase;
            }
            set
            {
                _rememberDatabase = value;
                Properties.Settings.Default.RememberDatabase = _rememberDatabase;
                Properties.Settings.Default.Save();
                OnPropertyChanged(propertyName: nameof(RememberDatabase));
            }
        }

        public ObservableCollection<Query> Queries
        {
            get
            {
                if (_queries == null)
                {
                    if (_dataService != null)
                    {
                        List<Query> queryList = _dataService.GetQueryList(_startDate, _endDate, out List<string> messages);
                        foreach (string message in messages)
                        {
                            Messages.Add(message);
                        }
                        _queries = new ObservableCollection<Query>(queryList);
                    }
                }

                return _queries;
            }
        }

        public List<Query> SelectedQueries
        {
            get
            {
                if (Queries != null)
                {
                    _selectedQueries = Queries.Where(query => (query.Selected == true)).ToList();
                }               
                return _selectedQueries;
            }
            set
            {
                _selectedQueries = value;
                OnPropertyChanged(propertyName: nameof(SelectedQueries));
            }
        }

        public string TextFileDirectory
        {
            get
            {
                return _textFileDirectory;
            }
            set
            {
                if (_textFileDirectory != value)
                {
                    _textFileDirectory = value;
                    Properties.Settings.Default.TextFileDirectory = _textFileDirectory;
                    Properties.Settings.Default.Save();
                    OnPropertyChanged(propertyName: nameof(TextFileDirectory));
                }
            }
        }

        public string ExcelFilePath
        {
            get
            {
                return _excelFilePath;
            }
            set
            {
                if (_excelFilePath != value)
                {
                    _excelFilePath = value;
                    Properties.Settings.Default.ExcelFilePath = _excelFilePath;
                    Properties.Settings.Default.Save();
                    OnPropertyChanged(propertyName: nameof(ExcelFilePath));
                }
            }
        }

        public DateTime StartDate
        {
            get
            {
                return _startDate;
            }
            set
            {
                if (_startDate != value)
                {
                    _startDate = value;                    

                    if (Queries != null)
                    {
                        foreach (Query query in Queries)
                        {
                            List<QueryParameter> startDateParams = query.QueryParameters.Entities.Where(parameter => (parameter.ParameterName == "[Start_Date]")).ToList();
                            foreach (QueryParameter qp in startDateParams)
                            {
                                qp.Value = value.ToShortDateString();
                            }
                        }
                    }

                    OnPropertyChanged(propertyName: nameof(StartDate));
                }
            }
        }

        public DateTime EndDate
        {
            get
            {
                return _endDate;
            }
            set
            {
                if (_endDate != value)
                {
                    _endDate = value;

                    if (Queries != null)
                    {
                        foreach (Query query in Queries)
                        {
                            List<QueryParameter> endDateParams = query.QueryParameters.Entities.Where(parameter => (parameter.ParameterName == "[End_Date]")).ToList();
                            foreach (QueryParameter qp in endDateParams)
                            {
                                qp.Value = value.ToShortDateString();
                            }
                        }
                    }

                    OnPropertyChanged(propertyName: nameof(EndDate));
                }
            }
        }

        public string ProcessTime
        {
            get
            {
                return _processTime;
            }
            set
            {
                if (_processTime != value)
                {
                    _processTime = value;
                    OnPropertyChanged(propertyName: nameof(ProcessTime));
                }
            }
        }

        public string StatusMessage
        {
            get
            {
                return _statusMessage;
            }
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged(propertyName: nameof(StatusMessage));
                }
            }
        }

        public bool Idle
        {
            get
            {
                return _idle;
            }
            set
            {
                _idle = value;
                OnPropertyChanged(propertyName: nameof(Idle));
            }
        }

        #endregion // Properties


        #region RelayCommand Properties -----------------------------------------------------------

        public RelayCommand BrowseForDatabaseCommand
        {
            get
            {
                if (_browseForDatabaseCommand == null)
                {
                    _browseForDatabaseCommand = new RelayCommand(objectParameter => BrowseForDatabase(), objectParameter => { return true; });
                }
                return _browseForDatabaseCommand;
            }
        }

        public RelayCommand BrowseForFileCommand
        {
            get
            {
                if (_browseForFileCommand == null)
                {
                    _browseForFileCommand = new RelayCommand(objectParameter => BrowseForExcelFile(), objectParameter => { return true; });
                }
                return _browseForFileCommand;
            }
        }

        public RelayCommand BrowseForDirectoryCommand
        {
            get
            {
                if (_browseForDirectoryCommand == null)
                {
                    _browseForDirectoryCommand = new RelayCommand(objectParameter => BrowseForDirectory(), objectParameter => { return true; });
                }
                return _browseForDirectoryCommand;
            }
        }

        public RelayCommand SetDateRangeCommand
        {
            get
            {
                if (_setDateRangeCommand == null)
                {
                    _setDateRangeCommand = new RelayCommand(objectParameter => SetDateRange(objectParameter), objectParameter => { return (_dataService != null); });
                }
                return _setDateRangeCommand;
            }
        }

        public RelayCommand ClearMessagesCommand
        {
            get
            {
                if (_clearMessagesCommand == null)
                {
                    _clearMessagesCommand = new RelayCommand(objectParameter => ClearMessages(), objectParameter => { return Messages.Count > 0; });
                }
                return _clearMessagesCommand;
            }
        }

        public RelayCommand OpenDirectoryCommand
        {
            get
            {
                if (_openDirectoryCommand == null)
                {
                    _openDirectoryCommand = new RelayCommand(objectParameter => { OpenDirectory(); }, objectParameter => { return true; });
                }
                return _openDirectoryCommand;
            }
        }

        public RelayCommand RunQueriesCommand
        {
            get
            {
                if (_runQueriesCommand == null)
                {
                    _runQueriesCommand = new RelayCommand(objectParameter =>
                    {
                        RunQueriesAsync();
                    },
                        objectParameter =>
                        {
                            if ((SelectedQueries == null) || (Idle == false)) return false;
                            return SelectedQueries.Count > 0;
                        });
                }
                return _runQueriesCommand;
            }
        }

        #endregion // RelayCommand Properties


        #region Command Functions -----------------------------------------------------------------

        private void SetStatus(string statusMessage, string processTime)
        {
            StatusMessage = statusMessage;
            ProcessTime = processTime;
        }

        private void BrowseForDatabase()
        {
            string savedDatabasePath = Properties.Settings.Default.DatabasePath;
            string databasePath = DatabasePath;
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            if (databasePath.Length > 0)
            {
                openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(databasePath);
                openFileDialog.FileName = System.IO.Path.GetFileName(databasePath);
            }
            else
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            openFileDialog.DefaultExt = ".accdb";
            openFileDialog.Filter = "Access Files (*.accdb)|*.accdb";
            openFileDialog.Multiselect = false;
            openFileDialog.CheckFileExists = true;
            if (openFileDialog.ShowDialog() == true)
            {
                databasePath = openFileDialog.FileName;
                if (!savedDatabasePath.Equals(databasePath, StringComparison.InvariantCultureIgnoreCase) && (RememberDatabase == true))
                {
                    Properties.Settings.Default.DatabasePath = databasePath;
                    Properties.Settings.Default.Save();
                }
                DatabasePath = databasePath;
            }
        }

        private void BrowseForExcelFile()
        {
            string savedExcelFilePath = Properties.Settings.Default.ExcelFilePath;
            string excelFilePath = ExcelFilePath;
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            if (!string.IsNullOrWhiteSpace(excelFilePath))
            {
                openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(excelFilePath);
                openFileDialog.FileName = System.IO.Path.GetFileName(excelFilePath);
            }
            else
            {
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileDialog.FileName = "results.xlsx";
            }
            openFileDialog.DefaultExt = ".xlsx";
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|Text Files (*.txt)|*.txt";
            openFileDialog.Multiselect = false;
            openFileDialog.CheckFileExists = false;
            if (openFileDialog.ShowDialog() == true)
            {
                excelFilePath = openFileDialog.FileName;
                ExcelFilePath = excelFilePath;
            }
        }

        private void BrowseForDirectory()
        {
            string savedTextFileDirectory = Properties.Settings.Default.TextFileDirectory;
            string textFileDirectory = TextFileDirectory;

            using (var browserDialog = new FolderBrowserDialog())
            {
                browserDialog.RootFolder = Environment.SpecialFolder.MyComputer;
                if (System.IO.Directory.Exists(textFileDirectory))
                {
                    browserDialog.SelectedPath = textFileDirectory;
                }
                browserDialog.Description = "Select Output Directory for Text Files";
                browserDialog.ShowNewFolderButton = true;
                DialogResult result = browserDialog.ShowDialog();

                if (result == DialogResult.OK)
                {
                    textFileDirectory = browserDialog.SelectedPath;
                    TextFileDirectory = textFileDirectory;
                }
            }
        }

        private void SetDateRange(object objectParameter)
        {
            string dateRange = objectParameter.ToString();

            switch (dateRange)
            {
                case "ThisWeek":
                    StartDate = DateTimeExtensions.StartOfWeek(DateTime.Today, DayOfWeek.Monday);
                    EndDate = DateTimeExtensions.EndOfWeek(DateTime.Today, DayOfWeek.Monday);
                    break;
                case "LastWeek":
                    StartDate = DateTimeExtensions.StartOfLastWeek(DateTime.Today, DayOfWeek.Monday);
                    EndDate = DateTimeExtensions.EndOfLastWeek(DateTime.Today, DayOfWeek.Monday);
                    break;
                case "CurrentMonth":
                    StartDate = DateTimeExtensions.FirstDayOfMonth(DateTime.Today);
                    EndDate = DateTimeExtensions.LastDayOfMonth(DateTime.Today);
                    break;
                default:
                    break;
            }
        }

        private void ClearMessages()
        {
            Messages.Clear();
        }

        private void OpenDirectory()
        {
            string directoryPath = string.Empty;

            if (TextOutput)
            {
                directoryPath = TextFileDirectory;
            }
            else
            {
                directoryPath = System.IO.Path.GetDirectoryName(ExcelFilePath);
            }

            if (System.IO.Directory.Exists(directoryPath))
            {
                try
                {
                    Process.Start("explorer.exe", directoryPath);
                }
                catch (Exception ex)
                {
                    Messages.Add(ex.Message);
                }
            }
            else
            {
                Messages.Add("Output directory not available.");
            }
        }

        private void RunQueriesAsync()
        {
            if ((SelectedQueries == null) || (SelectedQueries.Count <= 0))
            {
                return;
            }

            SetStatus("Running selected queries. Please wait.", string.Empty);

            string fullPath = string.Empty;
            string outputDirectory = string.Empty;
            string outputFile = string.Empty;
            string outputExtension = string.Empty;

            if (TextOutput)
            {
                fullPath = System.IO.Directory.Exists(TextFileDirectory) ? TextFileDirectory : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                TextFileDirectory = fullPath;
            }
            else
            {
                fullPath = ExcelFilePath;
                if (!string.IsNullOrWhiteSpace(fullPath))
                {
                    outputDirectory = System.IO.Path.GetDirectoryName(fullPath);
                    if (!System.IO.Directory.Exists(outputDirectory))
                    {
                        outputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    }

                    outputFile = System.IO.Path.GetFileNameWithoutExtension(fullPath);
                    outputExtension = System.IO.Path.GetExtension(fullPath);

                    if (outputExtension != ".xlsx")
                    {
                        outputExtension = ".xlsx";
                    }

                    if (!string.IsNullOrWhiteSpace(outputFile))
                    {
                        outputFile += outputExtension;
                    }
                    else
                    {
                        outputFile = "results.xlsx";
                    }
                }
                else
                {
                    outputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    outputFile = "results.xlsx";
                }

                fullPath = System.IO.Path.Combine(outputDirectory, outputFile);
                ExcelFilePath = fullPath;
            }

            var thisContext = SynchronizationContext.Current;

            Task.Factory.StartNew(delegate
            {
                Idle = false;
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                ExportResults(out List<string> messages);

                thisContext.Send((x) =>
                {
                    foreach (string message in messages)
                    {
                        Messages.Add(message);
                    }
                }, null);

                stopWatch.Stop();

                TimeSpan ts = stopWatch.Elapsed;
                SetStatus("Report generation complete.", String.Format("{0:00}:{1:00}.{2:00}", ts.Minutes, ts.Seconds, ts.Milliseconds / 10));

                Idle = true;
            });

        }

        private void ExportResults(out List<string> messages)
        {
            messages = new List<string>();

            if (TextOutput)
            {
                ExportToText(ref messages);
            }
            else
            {
                ExportToExcel(ref messages);
            }
        }


        private void ExportToText(ref List<string> messages)
        {
            string filePath = string.Empty;
            List<string> columnList = new List<string>();
            
            try
            {
                foreach (Query query in SelectedQueries)
                {
                    filePath = query.QueryName.Trim() + ".txt";
                    filePath = System.IO.Path.Combine(TextFileDirectory, filePath);

                    using (System.Data.DataTable dataTable = _dataService.GetResultsTable(query, out List<string> dataServiceMessages))
                    {
                        foreach (string message in dataServiceMessages)
                        {
                            messages.Add(message);
                        }

                        if (dataTable == null)
                        {
                            continue;
                        }

                        using (System.IO.StreamWriter file = new System.IO.StreamWriter(filePath))
                        {
                            foreach (System.Data.DataColumn column in dataTable.Columns)
                            {
                                columnList.Add(column.ColumnName);
                            }
                            file.WriteLine(String.Join("\t", columnList.ToArray()));
                            columnList.Clear();

                            int count = dataTable.Columns.Count;

                            foreach (System.Data.DataRow row in dataTable.Rows)
                            {
                                for (int i = 0; i < count; i++)
                                {
                                    columnList.Add(row[i].ToString());
                                }
                                file.WriteLine(String.Join("\t", columnList.ToArray()));
                                columnList.Clear();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                messages.Add(ex.Message);
            }
        }

        private void ExportToExcel(ref List<string> messages)
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Workbook wkbook = null;
            Workbooks wkbooks = null;

            bool fileExists = System.IO.File.Exists(ExcelFilePath);

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application
                {
                    DisplayAlerts = false
                };

                wkbooks = excel.Workbooks;

                if (fileExists)
                {
                    wkbook = wkbooks.Open(ExcelFilePath);
                }
                else
                {
                    wkbook = wkbooks.Add(Type.Missing);
                }
            }
            catch (Exception ex)
            {
                messages.Add(ex.Message);
            }

            if (excel == null)
            {
                return;
            }

            if (wkbook == null)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try
                {
                    excel.DisplayAlerts = true;
                    excel.Quit();
                }
                catch { }
                _dataService.ReleaseComObject(excel);
                return;
            }

            Sheets sheets = wkbook.Worksheets;
            Worksheet sheet = null;
            Range range = null;
            Range columns = null;
            Font font = null;

            string sheetName = string.Empty;

            // TODO: Devise a more sophisticated RegEx for transforming Query.QueryName values to spreadsheet names.
            // 31 character limit
            // Characters not permitted: \ / ? * [ ]
            //
            // Maybe:
            // ([^/\\?*\[\]]{1,31}|[A-Za-z0-9_]{1,31})

            Regex pattern = new Regex(@"[\\/\?*\[\]]");

            try
            {
                foreach (Query query in SelectedQueries)
                {
                    sheetName = query.QueryName;
                    sheetName = (sheetName.Length > 31) ? sheetName.Substring(0, 30) : sheetName;
                    sheetName = pattern.Replace(sheetName, "-");
                    sheetName.Trim();

                    for (int x = 1; x <= sheets.Count; x++)
                    {
                        sheet = (Worksheet)sheets[x];
                        if (sheet.Name == sheetName)
                        {
                            sheet.Delete();
                            break;
                        }

                        _dataService.ReleaseComObject(sheet);

                        sheet = null;
                    }

                    sheet = sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    sheet.Name = sheetName;

                    using (System.Data.DataTable dataTable = _dataService.GetResultsTable(query, out List<string> dataServiceMessages))
                    {
                        foreach (string message in dataServiceMessages)
                        {
                            messages.Add(message);
                        }

                        if (dataTable == null)
                        {
                            continue;
                        }

                        int columnCount = dataTable.Columns.Count;
                        int rowCount = dataTable.Rows.Count;

                        for (int c = 0; c < columnCount; c++)
                        {
                            range = (Range)sheet.Cells[1, c + 1];
                            range.Value2 = dataTable.Columns[c].ColumnName;
                            font = range.Font;
                            font.Bold = true;
                        }

                        for (int r = 0; r < rowCount; r++)
                        {
                            for (int c = 0; c < columnCount; c++)
                            {
                                range = (Range)sheet.Cells[r + 2, c + 1];
                                range.Value2 = dataTable.Rows[r][c].ToString();
                            }
                        }
                        columns = sheet.Columns;
                        columns.AutoFit();
                    }
                }

                if (fileExists)
                {
                    wkbook.Save();
                }
                else
                {
                    wkbook.SaveAs(ExcelFilePath);
                }
            }
            catch (Exception ex)
            {
                messages.Add(ex.Message);
            }

            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            _dataService.ReleaseComObject(range);
            _dataService.ReleaseComObject(columns);
            _dataService.ReleaseComObject(font);
            _dataService.ReleaseComObject(sheet);
            _dataService.ReleaseComObject(sheets);

            wkbook.Close(Type.Missing, Type.Missing, Type.Missing);
            _dataService.ReleaseComObject(wkbook);

            excel.DisplayAlerts = true;
            excel.Quit();
            _dataService.ReleaseComObject(excel);
        }

        #endregion // Command Functions
    }
}

