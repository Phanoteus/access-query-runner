using System.Data.OleDb;

namespace QueryRunner.Data.Entities
{
    public class QueryParameter : EntityBase
    {
        private string _parameterName;
        private OleDbType _type = OleDbType.Empty;
        private dynamic _value;

        public QueryParameter() { }
        public QueryParameter(string parameterName, OleDbType type, dynamic value)
        {
            _parameterName = parameterName;
            _type = type;
            _value = value;
        }

        public string ParameterName
        {
            get { return _parameterName; }
            set
            {
                if (_parameterName == value) return;

                _parameterName = value;
                OnPropertyChanged(nameof(ParameterName));
            }
        }

        public OleDbType Type
        {
            get { return _type; }
            set
            {
                if (_type == value) return;

                _type = value;
                OnPropertyChanged(nameof(Type));
            }
        }

        public dynamic Value
        {
            get { return _value; }
            set
            {
                _value = value;
                OnPropertyChanged(nameof(Value));
            }
        }
    }
}
