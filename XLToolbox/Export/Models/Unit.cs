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
            Dictionary<Unit, double> inner;
            double factor;
            if (ConversionTable.TryGetValue(unit, out inner) &&
                inner.TryGetValue(targetUnit, out factor))
            {
                return fromValue * factor;
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

        private static Dictionary<Unit, Dictionary<Unit, double>> ConversionTable
        {
            get
            {
                if (_conversionTable == null)
                {
                    _conversionTable = new Dictionary<Unit,Dictionary<Unit,double>>()
                    {
                        { Unit.Inch, new Dictionary<Unit, double>()
                            {
                                { Unit.Inch, 1.0 },
                                { Unit.Millimeter, 25.4 },
                                { Unit.Point, 72.0 }
                            }
                        },
                        { Unit.Millimeter, new Dictionary<Unit, double>()
                            {
                                { Unit.Inch, 1.0/25.4 },
                                { Unit.Millimeter, 1.0 },
                                { Unit.Point, 72.0/25.4 }
                            }
                        },
                        { Unit.Point, new Dictionary<Unit, double>()
                            {
                                { Unit.Inch, 1.0/72.0 },
                                { Unit.Millimeter, 25.4/72.0 },
                                { Unit.Point, 1.0 }
                            }
                        }
                    };
                }
                return _conversionTable;
            }
        }

        #endregion

        #region Private static fields

        private static Dictionary<Unit, Dictionary<Unit, double>> _conversionTable;

        #endregion

    }
}
