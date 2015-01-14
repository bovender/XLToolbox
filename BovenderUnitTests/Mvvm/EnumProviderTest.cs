using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using Bovender.Mvvm;

namespace Bovender.UnitTests.Mvvm
{
    enum TestEnum
    {
        one,
        two,
        [System.ComponentModel.Description("drei")]
        three
    }

    [TestFixture]
    class EnumProviderTest
    {
        [Test]
        public void Choices()
        {
            EnumProvider<TestEnum> provider = new EnumProvider<TestEnum>();
            Assert.AreEqual(Enum.GetNames(typeof(TestEnum)).Length, provider.Choices.Count(),
                "Choices array has incorrect length.");

            Assert.AreEqual("two", provider.Choices.ToList()[1].ToString());
            Assert.AreEqual("drei", provider.Choices.ToList()[2].ToString());
        }

        [Test]
        public void EnumToString()
        {
            EnumProvider<TestEnum> provider = new EnumProvider<TestEnum>();
            provider.AsEnum = TestEnum.three;
            Assert.AreEqual("drei", provider.SelectedItem.ToString());
            provider.AsEnum = TestEnum.two;
            Assert.AreEqual("two", provider.SelectedItem.ToString());
        }

        /*
        [Test]
        public void StringToEnum()
        {
            EnumProvider<TestEnum> provider = new EnumProvider<TestEnum>();
            provider.SelectedItem = "one";
            Assert.AreEqual(TestEnum.one, provider.AsEnum);
            provider.SelectedItem = "drei";
            Assert.AreEqual(TestEnum.three, provider.AsEnum);
        }
         */
    }
}
