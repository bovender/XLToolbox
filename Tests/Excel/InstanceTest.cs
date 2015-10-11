/* InstanceTest.cs
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
using System.Linq;
using System.Text;
using NUnit.Framework;
using XLToolbox.Excel.ViewModels;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Test.Excel
{
    [TestFixture]
    class InstanceTest
    {
        [Test]
        public void CountOpenWorkbooks()
        {
            Instance i = Instance.Default;
            i.CreateWorkbook();
            i.CreateWorkbook();
            Assert.AreEqual(i.Application.Workbooks.Count, i.CountOpenWorkbooks);
        }

        [Test]
        public void CountUnsavedWorkbooks()
        {
            Instance i = Instance.Default;
            for (int n = 0; n < 5; n++)
            {
                i.CreateWorkbook();
            }
            foreach (Workbook w in i.Application.Workbooks)
            {
                w.Saved = true;
            }
            i.Application.Workbooks[2].Saved = false;
            i.Application.Workbooks[3].Saved = false;
            Assert.AreEqual(2, i.CountUnsavedWorkbooks);
        }

        [Test]
        public void CountSavedWorkbooks()
        {
            Instance i = Instance.Default;
            for (int n = 0; n < 5; n++)
            {
                i.CreateWorkbook();
            }
            foreach (Workbook w in i.Application.Workbooks)
            {
                w.Saved = false;
            }
            i.Application.Workbooks[1].Saved = true;
            i.Application.Workbooks[2].Saved = true;
            i.Application.Workbooks[3].Saved = true;
            Assert.AreEqual(3, i.CountSavedWorkbooks);
        }
    }
}
