using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;
using System.ComponentModel;

namespace Bovender.Mvvm.ViewModels
{
    /// <summary>
    /// Collection of view models that automatically syncs with an
    /// associated collection of model objects.
    /// </summary>
    /// <typeparam name="TModel"></typeparam>
    /// <typeparam name="TViewModel"></typeparam>
    public abstract class ViewModelCollection<TModel, TViewModel>
        : ObservableCollection<TViewModel> where TViewModel : ViewModelBase
    {
        #region Abstract methods

        protected abstract TViewModel CreateViewModel(TModel model);

        #endregion

        #region Constructor

        public ViewModelCollection(ObservableCollection<TModel> modelCollection)
        {
            _modelCollection = modelCollection;
            _modelCollection.CollectionChanged += _modelCollection_CollectionChanged;
            this.CollectionChanged += ViewModelCollection_CollectionChanged;
            BuildViewModelCollection();
        }

        #endregion

        #region Event handlers

        /// <summary>
        /// Propagates changes in the view model collection to the model collection.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ViewModelCollection_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (_working) return;
        
            _working = true;
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    foreach (TViewModel vm in e.NewItems)
                    {
                        _modelCollection.Add((TModel)vm.RevealModelObject());
                    }
                    break;
                case NotifyCollectionChangedAction.Remove:
                    foreach (TViewModel vm in e.OldItems)
                    {
                        _modelCollection.Remove((TModel)vm.RevealModelObject());
                    }
                    break;
            }
            _working = false;
        }

        /// <summary>
        /// Propagates changes in the model collection to this collection of view models.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _modelCollection_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (_working) return;

            _working = true;
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    foreach (TModel m in e.NewItems)
                    {
                        Add(CreateViewModel(m));
                    }
                    break;
                case NotifyCollectionChangedAction.Remove:
                    foreach (TModel m in e.OldItems)
                    {
                        Items.Remove(
                            Items.FirstOrDefault(
                                (TViewModel vm) => vm.IsViewModelOf(m)
                            )
                        );
                    }
                    break;
            }
            _working = false;
        }

        #endregion

        #region Private methods

        protected void BuildViewModelCollection()
        {
            if (_working) return;

            _working = true;
            try
            {
                this.Clear();
                foreach (TModel m in _modelCollection)
                {
                    this.Add(CreateViewModel(m));
                }
            }
            finally
            {
                _working = false;
            }
        }

        #endregion

        #region Private fields

        readonly ObservableCollection<TModel> _modelCollection;
        bool _working;

        #endregion
    }
}
