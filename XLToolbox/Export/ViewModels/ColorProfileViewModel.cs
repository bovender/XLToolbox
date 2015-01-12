/* ColorProfileViewModel.cs
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
using Bovender.Mvvm.ViewModels;
using Bovender.Unmanaged;
using XLToolbox.Export.Models;
using System.Runtime.InteropServices;
using XLToolbox.Export.Lcms;

namespace XLToolbox.Export.ViewModels
{
    using cmsHProfile = IntPtr;
    using cmsBool = Boolean;

    /// <summary>
    /// View model for an ICS color profile.
    /// </summary>
    [Serializable]
    public class ColorProfileViewModel : ViewModelBase
    {
        #region Factory

        /// <summary>
        /// Creates a new ColorProfileViewModel from a given file, but only
        /// if the color profile's color space is known by the view model.
        /// </summary>
        /// <param name="fileName">Full file name of the ICS color profile.</param>
        /// <returns>Instance of ColorProfileViewModel, or null if the color
        /// profile has an unknown color space.</returns>
        public static ColorProfileViewModel CreateFromFile(string fileName)
        {
            cmsHProfile h = cmsOpenProfileFromFile(fileName, "r");
            ColorProfileViewModel cpvm = null;
            if (h != IntPtr.Zero)
            {
                ColorSpaceSignature c;
                try
                {
                    // Only create a view model if we can handle
                    // the current profile's color space.
                    c = cmsGetColorSpace(h);
                    cpvm = new ColorProfileViewModel(fileName);
                    cpvm.ColorSpace = c.ToColorSpace();
                }
                finally
                {
                    cmsCloseProfile(h);
                }
            }
            return cpvm;
        }

        #endregion

        #region Public properties

        public ColorSpace ColorSpace { get; private set; }

        public string Name
        {
            get
            {
                return _fileName;
            }
        }

        #endregion

        #region Protected constructor

        /// <summary>
        /// Instantiates the view model given a color profile's file name.
        /// </summary>
        /// <param name="fileName">Complete file name to the color profile.</param>
        /// <exception cref="ArgumentException">if the color profile does not
        /// exist or is not accessible.</exception>
        /// <remarks>
        /// The constructor is protected in order to enforce the factory method,
        /// which takes care of restricting instantiation to profiles with color
        /// spaces that we can actually handle.
        /// </remarks>
        protected ColorProfileViewModel(string fileName)
            :base()
        {
            _fileName = System.IO.Path.GetFileName(fileName);
        }

        #endregion

        #region Overrides

        public override string DisplayString
        {
            get
            {
                return _fileName;
            }
        }

        public override string ToString()
        {
            return DisplayString;
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            // TODO: Reveal color profile model object
            throw new NotImplementedException();
        }

        #endregion

        #region Private fields

        /// <summary>
        /// File name of the color profile (without path).
        /// </summary>
        private string _fileName;

        #endregion

        #region P/Invokes

        [DllImport("lcms2.dll", EntryPoint = "cmsOpenProfileFromFile")]
        private static extern cmsHProfile cmsOpenProfileFromFile(string iccProfile, string access);

        [DllImport("lcms2.dll", EntryPoint = "cmsCloseProfile")]
        private static extern cmsBool cmsCloseProfile(cmsHProfile hProfile);

        [DllImport("lcms2.dll", EntryPoint = "cmsGetColorSpace")]
        private static extern ColorSpaceSignature cmsGetColorSpace(cmsHProfile hProfile);

        #endregion
    }
}
