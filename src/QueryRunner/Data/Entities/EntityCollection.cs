using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueryRunner.Data.Entities
{
    public class EntityCollection<T> : EntityBase where T : INotifyPropertyChanged
    {
        private bool _entityAdded = false;
        private bool _entityDeleted = false;

        public EntityCollection(List<T> list = null)
        {
            if (list != null)
            {
                Entities = new ObservableEntityCollection<T>(list);
            }
            else
            {
                Entities = new ObservableEntityCollection<T>();
            }

            Entities.CollectionChanged += (sender, e) =>
            {
                if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add)
                {
                    EntityAdded = true;
                }
                if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Remove)
                {
                    EntityDeleted = true;
                }
                OnPropertyChanged(nameof(Entities));
            };
            Entities.ItemPropertyChanged += (sender, e) =>
            {
                OnPropertyChanged(nameof(Entities));
            };
        }

        public ObservableEntityCollection<T> Entities
        {
            get;
            private set;
        }

        public bool EntityAdded
        {
            get
            {
                return _entityAdded;
            }
            set
            {
                if (_entityAdded != value)
                {
                    _entityAdded = value;
                    OnPropertyChanged(nameof(EntityAdded));
                }
            }
        }

        public bool EntityDeleted
        {
            get
            {
                return _entityDeleted;
            }
            set
            {
                if (_entityDeleted != value)
                {
                    _entityDeleted = value;
                    OnPropertyChanged(nameof(EntityDeleted));
                }
            }
        }
    }
}
