/* Jumper.cs
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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Extensions;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Excel.Models;
using System.Text.RegularExpressions;

namespace XLToolbox
{
    public class Jumper
    {
        #region Factory

        public static Jumper FromSelection()
        {
            Jumper jumper = null;
            var selection = Instance.Default.Application.Selection;
            Range r = selection as Range;
            if (r != null)
            {
                string s = String.Format(r.Value2);
                Logger.Info("FromSelection: {0}", s);
                jumper = new Jumper(s);
            }
            else
            {
                Logger.Info("FromSelection: Cannot create instance, selection is a {0}",
                    selection.GetType().FullName);
            }
            Bovender.ComHelpers.ReleaseComObject(selection);
            return jumper;
        }

        #endregion

        #region Properties

        public string Target
        {
            get
            {
                return _target;
            }
            set
            {
                _target = value;
                Parse();
            }
        }

        public Uri Uri { get; private set; }

        public Reference Reference { get; private set; }

        public bool CanJump { get; private set; }

        public bool IsReference { get; private set; }

        public bool IsFile { get; private set; }

        public bool IsWebUrl { get; private set; }

        #endregion

        #region Constructor

        public Jumper(string target)
        {
            Target = target;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Jumps to the target.
        /// </summary>
        /// <returns>True on success, false on failure</returns>
        public bool Jump()
        {
            if (CanJump)
            {
                bool success = false;
                if (IsReference)
                {
                    Logger.Info("Jump: Reference: {0}", Reference.ReferenceString);
                    Reference.Activate();
                    success = true;
                }
                else if (IsFile)
                {
                    string path = Uri.LocalPath;
                    Logger.Info("Jump: Path: {0}", path);
                    Workbook workbook = null;
                    if (_workbookPattern.Value.IsMatch(path))
                    {
                        Logger.Info("Jump: Attempting to open workbook");
                        workbook = Instance.Default.LocateWorkbook(path);
                        success = workbook != null;
                        if (success)
                        {
                            ((_Workbook)workbook).Activate();
                            Logger.Info("Jump: Workbook opened");
                        }
                    }
                    if (!success)
                    {
                        if (System.IO.File.Exists(path) || System.IO.Directory.Exists(path))
                        {
                            Logger.Info("Jump: Starting process from path");
                            System.Diagnostics.Process.Start(path);
                            success = true;
                        }
                        else
                        {
                            Logger.Warn("Jump: Unable to locate target \"{0}\"", path);
                        }
                    }
                    Bovender.ComHelpers.ReleaseComObject(workbook);
                }
                else if (IsWebUrl)
                {
                    string url = Uri.AbsoluteUri;
                    Logger.Info("Jump: URL: {0}", url);
                    System.Diagnostics.Process.Start(url);
                    success = true;
                }
                return success;
            }
            else
            {
                Logger.Warn("Jump: Cannot jump to this target: \"{0}\"", Target);
                return false;
            }
        }

        #endregion

        #region Private methods

        private void Parse()
        {
            Logger.Info("Parse: Target: \"{0}\"", Target);
            Reference = new Excel.Models.Reference(Target);
            IsReference = Reference.IsValid;
            if (!String.IsNullOrEmpty(Target))
            {
                if (_genericUrlPattern.Value.IsMatch(Target))
                {
                    Logger.Info("Parse: Detected generic URL pattern, prepending http://");
                    _target = "http://" + _target;
                }
                try
                {
                    Uri = new Uri(Target);
                    IsFile = !IsReference && Uri.IsFile;
                    IsWebUrl = !IsReference && Uri.IsWellFormedOriginalString() && !Uri.IsFile;
                }
                catch (Exception e)
                {
                    Logger.Warn("Parse: Failed to create Uri object");
                    Logger.Warn(e);
                }
            }
            else
            {
                IsFile = false;
                IsWebUrl = false;
                Uri = null;
            }
            CanJump = (IsReference || IsFile || IsWebUrl);
            Logger.Debug("Parse: IsReference: {0}, IsFile: {1}, IsWebUrl: {2}, CanJump: {3}",
                IsReference, IsFile, IsWebUrl, CanJump);
        } 

        #endregion

        #region Private fields
	
        private string _target;
        private static readonly Lazy<Regex> _workbookPattern = new Lazy<Regex>(
            () => new Regex(@"\.xl(s|t|sx|sm|st)$"));
        private static readonly Lazy<Regex> _genericUrlPattern = new Lazy<Regex>(
            () => new Regex(@"^www\.[^.]+\."));

    	#endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}
