using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;
using NUnit.Framework;

namespace Bovender.UnitTests.Mvvm
{
    [TestFixture]
    public class ViewModelCollectionTest
    {
        ObservableCollection<TestModel> mc;
        ViewModelCollectionForTesting vmc;

        [SetUp]
        public void Setup()
        {
            mc = new ObservableCollection<TestModel>();
            vmc = new ViewModelCollectionForTesting(mc);
        }

        [Test]
        public void AddItemToModelCollection()
        {
            string testValue = "hello world";
            mc.Add(new TestModel(testValue));
            TestViewModel vm = vmc[0];
            Assert.AreEqual(testValue, vm.Value);
        }

        [Test]
        public void RemoveItemFromModelCollection()
        {
            TestModel m = new TestModel();
            mc.Add(m);
            Assert.AreEqual(1, vmc.Count);
            mc.Remove(m);
            Assert.AreEqual(0, vmc.Count);
        }

        [Test]
        public void AddItemToViewModelCollection()
        {
            string testValue = "hello world";
            TestModel m = new TestModel(testValue);
            TestViewModel vm = new TestViewModel(m);
            vmc.Add(vm);
            TestModel testm = mc[0];
            Assert.AreEqual(m.Value, testm.Value);
        }

        [Test]
        public void RemoveItemFromViewModelCollection()
        {
            TestViewModel vm = new TestViewModel(new TestModel());
            vmc.Add(vm);
            Assert.AreEqual(1, mc.Count);
            vmc.Remove(vm);
            ViewModelCollectionForTesting c = new ViewModelCollectionForTesting(mc);
            Assert.AreEqual(0, mc.Count);
        }
    }
}
