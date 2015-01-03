using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bovender.UnitTests.Mvvm
{
    class TestModel
    {
        public string Value { get; set; }

        public TestModel() { }

        public TestModel(string value)
            : this()
        {
            Value = value;
        }
    }
}
