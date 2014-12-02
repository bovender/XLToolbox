using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;
using System.ComponentModel;
using System.Collections;

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
            // The BuildViewModelCollection adds the event handlers
            // when done, so there is no need to add the event handlers
            // via SynchronizeOn() in the constructor. Avoid adding the
            // handlers twice...
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
            SynchronizeOff();
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    DoAddModelObjects(e.NewItems);
                    break;
                case NotifyCollectionChangedAction.Remove:
                    DoRemoveModelObjects(e.OldItems);
                    break;
                case NotifyCollectionChangedAction.Move:
                    // Don't do anything if items are moved.
                    break;
                case NotifyCollectionChangedAction.Replace:
                    DoRemoveModelObjects(e.OldItems);
                    DoAddModelObjects(e.NewItems);
                    break;
                case NotifyCollectionChangedAction.Reset:
                    BuildModelCollection();
                    break;
            }
            SynchronizeOn();
        }

        /// <summary>
        /// Propagates changes in the model collection to this collection of view models.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _modelCollection_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            SynchronizeOff();
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    DoAddViewModelObjects(e.NewItems);
                    break;
                case NotifyCollectionChangedAction.Remove:
                    DoRemoveViewModelObjects(e.OldItems);
                    break;
                case NotifyCollectionChangedAction.Move:
                    // No need to synchronize, we don't care about order
                    break;
                case NotifyCollectionChangedAction.Replace:
                    DoRemoveViewModelObjects(e.OldItems);
                    DoAddViewModelObjects(e.NewItems);
                    break;
                case NotifyCollectionChangedAction.Reset:
                    BuildViewModelCollection();
                    break;
            }
            SynchronizeOn();
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Turns on synchronization of the view model collection and the
        /// model collection by adding appropriate event handles.
        /// </summary>
        protected void SynchronizeOn()
        {
            _modelCollection.CollectionChanged += _modelCollection_CollectionChanged;
            this.CollectionChanged += ViewModelCollection_CollectionChanged;
        }

        /// <summary>
        /// Turns off synchronization of the view model collection and the
        /// model collection by removing the event handles.
        /// </summary>
        protected void SynchronizeOff()
        {
            _modelCollection.CollectionChanged -= _modelCollection_CollectionChanged;
            this.CollectionChanged -= ViewModelCollection_CollectionChanged;
        }

        protected void BuildViewModelCollection()
        {
            try
            {
                SynchronizeOff();
                this.Clear();
                foreach (TModel m in _modelCollection)
                {
                    this.Add(CreateViewModel(m));
                }
            }
            finally
            {
                SynchronizeOn();
            }
        }

        private void DoAddViewModelObjects(IList modelObjects)
        {
            foreach (TModel m in modelObjects)
            {
                Add(CreateViewModel(m));
            }
        }

        private void DoRemoveViewModelObjects(IList modelObjects)
        {
            foreach (TModel m in modelObjects)
            {
                Items.Remove(
                    Items.FirstOrDefault(
                        (TViewModel vm) => vm.IsViewModelOf(m)
                    )
                );
            }
        }

        private void BuildModelCollection()
        {
            try
            {
                SynchronizeOff();
                this.Clear();
                foreach (TViewModel vm in Items)
                {
                    _modelCollection.Add((TModel)vm.RevealModelObject());
                }
            }
            finally
            {
                SynchronizeOn();
            }
        }

        private void DoAddModelObjects(IList viewModelObjects)
        {
            foreach (TViewModel vm in viewModelObjects)
            {
                _modelCollection.Add((TModel)vm.RevealModelObject());
            }
        }

        private void DoRemoveModelObjects(IList viewModelObjects)
        {
            foreach (TViewModel vm in viewModelObjects)
            {
                _modelCollection.Remove((TModel)vm.RevealModelObject());
            }
        }


        #endregion

        #region Private fields

        readonly ObservableCollection<TModel> _modelCollection;

        #endregion
    }
}
