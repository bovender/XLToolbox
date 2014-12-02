using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;

namespace Bovender.UnitTests.Mvvm
{
    class TestViewModel : ViewModelBase
    {
        public string Value
        {
            get { return Model.Value; }
            set
            {
                Model.Value = value;
            }
        }

        TestModel Model { get; set; }

        public TestViewModel(TestModel model)
        {
            Model = model;
        }

        public override object RevealModelObject()
        {
            return Model;
        }
    }
}
