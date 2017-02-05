/* SelectionViewModel.cs
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
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Office.Interop.Excel;
using Windows = System.Windows;
using Bovender.Mvvm.ViewModels;
using Bovender.Extensions;
using XLToolbox.Excel.Models;

namespace XLToolbox.Excel.ViewModels
{
    /// <summary>
    /// View model for the current selection of a running
    /// Excel instance. Note that the selection wrapped by
    /// this view model always reflects the current selection
    /// of the Excel application; it is not fixed.
    /// </summary>
    public class SelectionViewModel : ViewModelBase, IDisposable
    {
        #region Events

        public event EventHandler<SelectionChangedEventArgs> SelectionChanged;

        #endregion

        #region Properties

        /// <summary>
        /// Exposes the bound application's Selection property.
        /// </summary>
        public dynamic Selection
        {
            [System.Diagnostics.DebuggerStepThrough]
            get
            {
                return _app.Selection;
            }
        }

        public Range Range
        {
            get
            {
                return Selection as Range;
            }
        }

        public bool IsRange
        {
            get
            {
                return Range != null;
            }
        }

        public Reference Reference
        {
            get
            {
                if (IsRange)
                {
                    return new Reference(Range);
                }
                else
                {
                    Logger.Warn("Reference: Selection is not a Range, returning null!");
                    return null;
                }
            }
        }

        public Windows.Rect Bounds
        {
            get
            {
                if (Selection == null)
                {
                    throw new InvalidOperationException(
                        "Cannot compute bounds of selection because nothing is selected in Excel.");
                }
                return ComputeBounds();
            }
        }

        public string Name
        {
            get
            {
                if (Selection != null)
                {
                    try
                    {
                        return Selection.Name;
                    }
                    catch (Exception)
                    {
                        Logger.Warn("Name: Accessing the Name property on the dynamic Selection throws an exception");
                        return null;
                    }
                }
                else
                {
                    Logger.Warn("Name: Null has no name!");
                    return null;
                }
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Copies the current selection to the clipboard. Returns true on success,
        /// false on failure.
        /// </summary>
        /// <returns></returns>
        public bool CopyToClipboard()
        {
            bool result = false;
            dynamic selection = _app.Selection;
            Logger.Info("CopyToClipboard: Selection is a '{0}'",
                Microsoft.VisualBasic.Information.TypeName(selection));
            try
            {
                selection.Copy();
                result = true;
            }
            catch (Exception e)
            {
                Logger.Warn("CopyToClipboard: An exception was raised by selection.Copy()");
                Logger.Warn(e);
                throw new Excel.CopyException(Strings.GraphicHasNoCopyMethod, e);
            }
            finally
            {
                Bovender.ComHelpers.ReleaseComObject(selection);
            }
            return result;
        }

        public void SaveToEmf(string fileName)
        {
            CopyToClipboard();
            // Clipboard data format is spelled "EnhancedMetafile" - case is important!
            // If case is incorrect, "invalid TYMED" exception will occur.
            Metafile emf = Windows.Clipboard.GetData("EnhancedMetafile") as Metafile;

        }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructs the view model by binding to a particular
        /// Excel application instance.
        /// </summary>
        /// <param name="excelApplication">Excel instance whose selection
        /// this view model wraps.</param>
        public SelectionViewModel(Application excelApplication)
            :base()
        {
            _app = excelApplication;
            if (_app != null)
            {
                _app.SheetActivate += Excel_SheetActivate;
                _app.WorkbookActivate += Excel_WorkbookActivate;
                _app.SheetSelectionChange += Excel_SelectionChange;
            }
            else
            {
                Logger.Warn("Constructor: excelApplication is null!");
            }
        }

        #endregion

        #region Disposal

        ~SelectionViewModel()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
                if (disposing)
                {
                    // Clean up managed resources
                }
            }
        }

        #endregion

        #region Event handlers

        void Excel_WorkbookActivate(Workbook Wb)
        {
            OnSelectionChanged();
        }

        void Excel_SheetActivate(object Sh)
        {
            OnSelectionChanged();
        }

        void Excel_SelectionChange(object Sh, Range Target)
        {
            OnSelectionChanged();
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return Selection;
        }

        #endregion

        #region Protected methods

        protected void OnSelectionChanged()
        {
            EventHandler<SelectionChangedEventArgs> h = SelectionChanged;
            if (h != null)
            {
                h(this, new SelectionChangedEventArgs(_app));
            }
        }

        #endregion

        #region Private methods

        private Windows.Rect ComputeBounds()
        {
            // An Excel selection can be a Range, an embedded ChartObject,
            // an embedded Shape, a Chart sheet or multiple ChartObjects or
            // Shapes. The Application.Selection may therefore be one of a number
            // of different objects. Unlike VBA, C#/VSTO does not know about
            // a DrawingOjects class, so that we must find out about the nature
            // of the Selection (single vs. multiple objects) by exclusion.
            Windows.Rect r = new Windows.Rect();
            try
            {
                r.Width = _app.Selection.Width;
                r.Height = _app.Selection.Height;
                r.Location = new Windows.Point(_app.Selection.Left, _app.Selection.Top);
            }
            catch
            {
                // Get the bounding rectangle of multiple objects.
                // LINQ would be more elegant here, but it does not seem to
                // work with the COM interop Selection object.
                // Excel's collections are 1-based!
                dynamic firstObject = _app.Selection[1];
                double left = firstObject.Left;
                double right = left + firstObject.Width;
                double top = firstObject.Top;
                double bottom = top + firstObject.Height;
                foreach (dynamic o in _app.Selection)
                {
                    if (o.Left < left) left = o.Left;
                    if (o.Top < top) top = o.Top;
                    if (o.Left + o.Width > right) right = o.Left + o.Width;
                    if (o.TOp + o.Height > bottom) bottom = o.Top + o.Height;
                }
                r.Location = new Windows.Point(left, top);
                r.Size = new Windows.Size(right-left, bottom-top);
            }
            return r;
        }

        #endregion

        #region Private fields

        private bool _disposed;
        private Application _app;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}
