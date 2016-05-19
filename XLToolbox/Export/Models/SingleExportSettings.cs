using Microsoft.Office.Interop.Excel;
/* SingleExportSettings.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2016 Daniel Kraus
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
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.Export.Models
{
    /// <summary>
    /// Holds settings for a specific single export process.
    /// </summary>
    public class SingleExportSettings : Settings
    {
        #region Factory

        /// <summary>
        /// Creates a single export settings for the current selection.
        /// </summary>
        public static SingleExportSettings CreateForSelection(Preset preset)
        {
            SelectionViewModel svm = new SelectionViewModel(Instance.Default.Application);
            // If the ActiveChart property of the Excel application is not null,
            // either a chart or 'something in the chart' is selected. To make sure
            // we don't attempt to export 'something in the chart', we select the
            // entire chart.
            // If there is no workbook open, accessing the ActiveChart property causes
            // a COM exception.
            object activeChart = null;
            try
            {
                activeChart = Instance.Default.Application.ActiveChart;
            }
            catch (System.Runtime.InteropServices.COMException) { }
            finally
            {
                if (activeChart != null)
                {
                    ChartViewModel cvm = new ChartViewModel(activeChart as Chart);
                    // Handle chart sheets and embedded charts differently
                    cvm.SelectSpecial();
                }
            }
            if (svm.Selection != null)
            {
                return new SingleExportSettings(preset, svm.Bounds.Width, svm.Bounds.Height,
                    UserSettings.UserSettings.Default.ExportUnit, true);
            }
            else
            {
                return new SingleExportSettings();
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Width of the selection in points.
        /// </summary>
        public double Width
        {
            get { return _width; }
            set
            {
                // Preserve aspect, but only if current dimensions are not 0
                if (PreserveAspect && _width != 0 && _height != 0)
                {

                    double _aspect = _height / _width;
                    _height = value * _aspect;
                }
                _width = value;
            }
        }

        /// <summary>
        /// Height of the selection in points.
        /// </summary>
        public double Height
        {
            get { return _height; }
            set
            {
                if (PreserveAspect && _width != 0 && _height != 0)
                {
                    double _aspect = _width / _height;
                    _width = value * _aspect;
                }
                _height = value;
            }
        }

        public bool PreserveAspect { get; set; }

        public Unit Unit
        {
            get
            {
                return _unit;
            }
            set
            {
                bool oldAspect = PreserveAspect;
                PreserveAspect = false;
                Height = _unit.ConvertTo(Height, value);
                Width = _unit.ConvertTo(Width, value);
                _unit = value;
                PreserveAspect = oldAspect;
            }
        }

        #endregion

        #region Constructors

        public SingleExportSettings()
            : base()
        {
            _unit = Models.Unit.Point;
            Preset = PresetsRepository.Default.First;
            PreserveAspect = true;
        }

        public SingleExportSettings(Preset preset, double width, double height, Unit unit, bool preserveAspect)
            : this()
        {
            Preset = preset;
            Width = width;
            Height = height;
            PreserveAspect = preserveAspect;
            _unit = unit;
        }

        #endregion

        #region Private fields

        double _width;
        double _height;
        Unit _unit;

        #endregion
    }
}
