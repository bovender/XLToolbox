/* StoreTest.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using NUnit.Framework;
using Microsoft.Office.Interop.Excel;
using XLToolbox.WorkbookStorage;
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.Test.WorkbookStorage
{
    [TestFixture]
    public class StoreTest
    {
        [Test]
        [ExpectedException(typeof(InvalidContextException))]
        public void InvalidContextCausesException()
        {
            Store storage = new Store();
            storage.Context = "made-up context that does not exist";
        }

        [Test]
        public void StoreInGlobalContext()
        {
            Store store = new Store();
            store.Context = "";
            store.Put("global setting", 1234);
            Assert.AreEqual(1234, store.Get("global setting", 0, 1, 2000));
        }

        [Test]
        [TestCase(333, 200, 400, 333)]
        [TestCase(123, 200, 400, 200)]
        [TestCase(423, 200, 400, 400)]
        public void StoreAndRetrieveIntegers(int i, int min, int max, int expect)
        {
            Store storage1 = new Store();
            string key = "Some property key";
            storage1.Put(key, i);
            storage1.Flush();
            Store storage2 = new Store();
            storage2.Context = storage1.Context;
            int j = storage2.Get(key, 0, min, max);
            Assert.AreEqual(expect, j);
        }

        [Test]
        public void StoreAndRetrieveString()
        {
            string s1 = "hello world";
            using (Store storage1 = new Store())
            {
                storage1.Put("my key", s1);
            }
            using (Store storage1 = new Store())
            {
                string s2 = storage1.Get("my key", "default");
                Assert.AreEqual(s1, s2);
            }
        }
        
        [Test]
        public void RetrieveNonExistingString()
        {
            Store store = new Store();
            string s = "hello world";
            Assert.AreEqual(s, store.Get("non-existing key", s));
        }

        [Test]
        [ExpectedException(typeof(EmptyKeyException))]
        public void DoNotAllowPutWithEmptyKey()
        {
            Store store = new Store();
            store.Put("", 123);
        }

        [Test]
        [ExpectedException(typeof(EmptyKeyException))]
        public void DoNotAllowGetWithEmptyKey()
        {
            Store store = new Store();
            string s = store.Get("", "not possible");
        }

        [Test]
        public void StoreRetrieveObject()
        {
            int number = 1234;
            ClassForTesting c = new ClassForTesting() { TestInt = number };
            Store store = new Store();
            string key = "testclass";
            store.Put<ClassForTesting>(key, c);
            ClassForTesting retrieved = store.Get<ClassForTesting>(key);
            Assert.NotNull(retrieved);
            Assert.AreEqual(number, retrieved.TestInt);
        }
    }
}
