/* ViewModelCollection.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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
        #region Public properties

        public int CountSelected
        {
            get { return _countSelected; }
            set
            {
                _countSelected = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CountSelected"));
            }
        }

        public TViewModel LastSelected
        {
            get { return _lastSelected; }
            set
            {
                _lastSelected = value;
                OnPropertyChanged(new PropertyChangedEventArgs("LastSelected"));
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Removes all selected view models from the collection.
        /// </summary>
        /// <remarks>
        /// Inspired by http://stackoverflow.com/a/26523396/270712
        /// </remarks>
        public void RemoveSelected()
        {
            var selected = Items.Where<TViewModel>((vm) => vm.IsSelected).ToList<TViewModel>();
            // Use Items.Remove() which does not trigger the CollectionChanged event.
            selected.ForEach((vm) => Items.Remove(vm));
            CountSelected = 0;
            LastSelected = null;
            OnPropertyChanged(new PropertyChangedEventArgs("Count"));
            OnPropertyChanged(new PropertyChangedEventArgs("Items[]"));
            OnCollectionChanged(new NotifyCollectionChangedEventArgs(
                NotifyCollectionChangedAction.Remove, selected
                )
            );
        }

        #endregion

        #region Events

        /// <summary>
        /// Relays property-changed events from the view models in the collection.
        /// </summary>
        public event EventHandler<PropertyChangedEventArgs> ViewModelPropertyChanged;

        #endregion

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

        private void viewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsSelected")
            {
                if (((TViewModel)sender).IsSelected)
                {
                    LastSelected = sender as TViewModel;
                    CountSelected++;
                }
                else
                {
                    LastSelected = null;
                    CountSelected--;
                }
            }
            OnViewModelPropertyChanged(sender, e);
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
                    AddViewModelWithEvent(m);
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
                AddViewModelWithEvent(m);
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
                vm.PropertyChanged += viewModel_PropertyChanged;
                _modelCollection.Add((TModel)vm.RevealModelObject());
            }
        }

        private void DoRemoveModelObjects(IList viewModelObjects)
        {
            foreach (TViewModel vm in viewModelObjects)
            {
                vm.PropertyChanged -= viewModel_PropertyChanged;
                _modelCollection.Remove((TModel)vm.RevealModelObject());
            }
        }

        private void AddViewModelWithEvent(TModel model)
        {
            TViewModel viewModel = CreateViewModel(model);
            viewModel.PropertyChanged += viewModel_PropertyChanged;
            Add(viewModel);
        }

        private void OnViewModelPropertyChanged(object sender, PropertyChangedEventArgs args)
        {
            if (ViewModelPropertyChanged != null)
            {
                ViewModelPropertyChanged(sender, args);
            }
        }

        #endregion

        #region Private fields

        readonly ObservableCollection<TModel> _modelCollection;
        private int _countSelected;
        private TViewModel _lastSelected;

        #endregion
    }
}
