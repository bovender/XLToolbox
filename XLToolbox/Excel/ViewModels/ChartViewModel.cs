/* ChartViewModel.cs
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
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace XLToolbox.Excel.ViewModels
{
    class ChartViewModel : Bovender.Mvvm.ViewModels.ViewModelBase
    {
        #region Public properties

        public Chart Chart
        {
            [DebuggerStepThrough]
            get
            {
                return _chart;
            }
            set
            {
                _chart = value;
                var parent = _chart.Parent;
                IsEmbedded = (parent is ChartObject);
                if (Marshal.IsComObject(parent)) Marshal.ReleaseComObject(parent);
                OnPropertyChanged("Chart");
                OnPropertyChanged("IsEmbedded");
            }
        }

        public bool IsEmbedded { get; protected set; }

        #endregion

        #region Public methods

        /// <summary>
        /// Selects either the Chart if it is a chart sheet, or the parent
        /// ChartObject if it is an embedded chart.
        /// </summary>
        public void SelectSpecial()
        {
            if (Chart == null)
            {
                throw new ArgumentNullException("Cannot select chart because chart is null");
            }
            if (IsEmbedded)
            {
                var parent = Chart.Parent;
                parent.Select();
                if (Marshal.IsComObject(parent)) Marshal.ReleaseComObject(parent);
            }
            else
            {
                ((_Chart)Chart).Select();
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructs a view model without model.
        /// </summary>
        public ChartViewModel() : base() { }

        /// <summary>
        /// Constructs a new view model for the given <paramref name="chart"/>.
        /// </summary>
        /// <param name="chart">Chart model.</param>
        public ChartViewModel(Chart chart)
            : this()
        {
            Chart = chart;
        }

        #endregion

        #region ViewModelBase overrides

        public override object RevealModelObject()
        {
            return _chart;
        }

        #endregion

        #region Private fields

        Chart _chart;

        #endregion
    }
}
