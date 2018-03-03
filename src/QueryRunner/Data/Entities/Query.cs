using System.Data;

namespace QueryRunner.Data.Entities
{
    public class Query : EntityBase
    {
        private string _queryName;
        private string _queryDefinition;
        private CommandType _queryCommandType;
        private EntityCollection<QueryParameter> _queryParameters;
        private bool _selected = false;
        private bool _valid = true;

        public Query()
        {
            _queryParameters = new EntityCollection<QueryParameter>();
        }

        public string QueryName
        {
            get { return _queryName; }
            set
            {
                if (_queryName == value) return;

                _queryName = value;
                OnPropertyChanged(nameof(QueryName));
            }
        }

        public string QueryDefinition
        {
            get { return _queryDefinition; }
            set
            {
                if (_queryDefinition == value) return;

                _queryDefinition = value;
                OnPropertyChanged(nameof(QueryDefinition));
            }
        }

        public CommandType QueryCommandType
        {
            get { return _queryCommandType; }
            set
            {
                if (_queryCommandType == value) return;

                _queryCommandType = value;
                OnPropertyChanged(nameof(QueryCommandType));
            }
        }

        public EntityCollection<QueryParameter> QueryParameters
        {
            get { return _queryParameters; }
            set
            {
                _queryParameters = value;
                OnPropertyChanged(nameof(QueryParameters));
            }
        }

        public bool Selected
        {
            get { return _selected; }
            set
            {
                _selected = value;
                OnPropertyChanged(nameof(Selected));
            }
        }

        public bool Valid
        {
            get { return _valid; }
            set
            {
                _valid = value;
                OnPropertyChanged(nameof(Valid));
            }
        }

    }
}
