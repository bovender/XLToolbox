/* Helpers.cs
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
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Test
{
    static class Helpers
    {
        public static void CreateSomeCharts(Worksheet worksheet, int number)
        {
            Random random = new Random(worksheet.GetHashCode());
            ChartObjects cos = worksheet.ChartObjects();
            ChartObject co;
            SeriesCollection sc;
            for (int n = 0; n < number; n++)
            {
                co = cos.Add(10 + random.Next(30), 10 + random.Next(40),
                    300 + random.Next(100), 200 + random.Next(50));
                sc = co.Chart.SeriesCollection();
                sc.Add(worksheet.Range["A1:A4"]);
            }
        }

        public static void CreateSomeShapes(Worksheet worksheet, int number)
        {
            Random random = new Random(worksheet.GetHashCode());
            Shapes shapes = worksheet.Shapes;
            for (int n = 0; n < number; n++)
            {
                shapes.AddLine(10 + random.Next(30), 10 + random.Next(40),
                    300 + random.Next(100), 200 + random.Next(50));
            }
        }
    }
}
