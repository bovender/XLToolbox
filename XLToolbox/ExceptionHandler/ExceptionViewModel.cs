/* ExceptionViewModel.cs
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
using System.Collections.Specialized;
using System.Reflection;
using Bovender.Unmanaged;
using XLToolbox.Versioning;

namespace XLToolbox.ExceptionHandler
{
    public class ExceptionViewModel : Bovender.ExceptionHandler.ExceptionViewModel
    {
        #region Additional properties for the exception report

        public string ExcelVersion
        {
            get
            {
                return Excel.ViewModels.Instance.Default.HumanFriendlyVersion;
            }
        }

        public string ExcelBitness
        {
            get
            {
                return Environment.Is64BitProcess ? "64-bit" : "32-bit";
            }
        }

        public string ToolboxVersion
        {
            get
            {
                return SemanticVersion.CurrentVersion().ToString();
            }
        }

        public string FreeImageVersion
        {
            get
            {
                try
                {
                    using (DllManager dllMan = new DllManager())
                    {
                        dllMan.LoadDll("FreeImage.dll");
                        if (FreeImageAPI.FreeImage.IsAvailable())
                        {
                            return FreeImageAPI.FreeImage.GetVersion();
                        }
                        else
                        {
                            return "Not available";
                        }
                    }
                }
                catch (Exception e)
                {
                    return e.Message;
                }
            }
        }

        #endregion

        #region constructor

        public ExceptionViewModel(Exception e) : base(e) { }

        #endregion

        #region Overrides

        public override object RevealModelObject()
        {
            return Exception;
        }

        protected override NameValueCollection GetPostValues()
        {
            NameValueCollection v = base.GetPostValues();
            v["excel_version"] = ExcelVersion;
            v["excel_bitness"] = ProcessBitness;
            v["freeimage_version"] = FreeImageVersion;
            v["toolbox_version"] = ToolboxVersion;
            return v;
        }

        protected override Uri GetPostUri()
        {
            return new Uri(Properties.Settings.Default.ExceptionPostUrl);
        }

        protected override string DevPath()
        {
            return @"x:\Code\xltoolbox\NG\";
        }

        #endregion
    }
}
