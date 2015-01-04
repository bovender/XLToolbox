/* Unit.cs
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
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace XLToolbox.Export.Models
{
    public enum Unit
    {
        [Description("pt")]
        Point,
        [Description("in")]
        Inch,
        [Description("mm")]
        Millimeter
    }

    public static class UnitsExtensions
    {
        #region Extension methods

        public static double ConvertTo(this Unit unit, double fromValue, Unit targetUnit)
        {
            Dictionary<Unit, decimal> inner;
            decimal factor;
            if (ConversionTable.TryGetValue(unit, out inner) &&
                inner.TryGetValue(targetUnit, out factor))
            {
                return Convert.ToDouble(factor * (decimal)fromValue);
            }
            else
            {
                throw new NotImplementedException(String.Format(
                    "No factor defined for conversion from {0} to {1}.",
                    unit.ToString(), targetUnit.ToString()));
            }
        }

        #endregion

        #region Private static properties

        private static Dictionary<Unit, Dictionary<Unit, decimal>> ConversionTable
        {
            get
            {
                if (_conversionTable == null)
                {
                    _conversionTable = new Dictionary<Unit,Dictionary<Unit,decimal>>()
                    {
                        { Unit.Inch, new Dictionary<Unit, decimal>()
                            {
                                { Unit.Inch, 1.0m },
                                { Unit.Millimeter, 25.4m },
                                { Unit.Point, 72.0m }
                            }
                        },
                        { Unit.Millimeter, new Dictionary<Unit, decimal>()
                            {
                                { Unit.Inch, 1.0m/25.4m },
                                { Unit.Millimeter, 1.0m },
                                { Unit.Point, 72.0m/25.4m }
                            }
                        },
                        { Unit.Point, new Dictionary<Unit, decimal>()
                            {
                                { Unit.Inch, 1.0m/72.0m },
                                { Unit.Millimeter, 25.4m/72.0m },
                                { Unit.Point, 1.0m }
                            }
                        }
                    };
                }
                return _conversionTable;
            }
        }

        #endregion

        #region Private static fields

        private static Dictionary<Unit, Dictionary<Unit, decimal>> _conversionTable;

        #endregion

    }
}
