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
