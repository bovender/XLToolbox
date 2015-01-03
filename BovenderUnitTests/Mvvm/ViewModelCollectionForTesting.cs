using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;
using Bovender.Mvvm.ViewModels;

namespace Bovender.UnitTests.Mvvm
{
    class ViewModelCollectionForTesting : ViewModelCollection<TestModel, TestViewModel>
    {
        public ViewModelCollectionForTesting(ObservableCollection<TestModel> modelCollection)
            : base(modelCollection)
        {
        }

        protected override TestViewModel CreateViewModel(TestModel model)
        {
            return new TestViewModel(model);
        }
    }
}
