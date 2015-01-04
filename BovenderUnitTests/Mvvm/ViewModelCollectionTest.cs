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
            Assert.AreEqual(1, mc.Count);
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

        [Test]
        public void SelectViewModels()
        {
            int n = 10;
            int s = 3;
            for (int i = 0; i < n; i++)
            {
                vmc.Add(new TestViewModel(new TestModel()));
            }
            // Select 's' number of view model objects. There's probably a more elegant way to do this.
            vmc[1].IsSelected = true;
            Assert.True(vmc[1].Equals(vmc.LastSelected),
                "View model at index 1 should be the LastSelected view model but isn't.");
            vmc[4].IsSelected = true;
            Assert.True(vmc[4].Equals(vmc.LastSelected),
                "View model at index 4 should be the LastSelected view model but isn't.");
            vmc[5].IsSelected = true;
            Assert.True(vmc[5].Equals(vmc.LastSelected),
                "View model at index 5 should be the LastSelected view model but isn't.");
            Assert.AreEqual(s, vmc.CountSelected, "Incorrect number of selected view models.");
            vmc.RemoveSelected();
            Assert.AreEqual(0, vmc.CountSelected,
                "After deleting selected view models, CountSelected should be 0.");
            Assert.AreEqual(n - s, vmc.Count,
                String.Format("There should be only {0} *view model* objects left.", n - s));
            Assert.AreEqual(n - s, mc.Count,
                String.Format("There should be only {0} *model* objects left.", n - s));
        }
    }
}
