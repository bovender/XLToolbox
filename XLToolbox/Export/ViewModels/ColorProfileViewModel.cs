/* ColorProfileViewModel.cs
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
using Bovender.Mvvm.ViewModels;
using Bovender.Unmanaged;
using XLToolbox.Export.Models;
using System.Runtime.InteropServices;
using XLToolbox.Export.Lcms;

namespace XLToolbox.Export.ViewModels
{
    using cmsHProfile = IntPtr;
    using cmsHTransform = IntPtr;
    using cmsUInt32Number = UInt32;
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
        /// Note that the lcms2.dll must have been loaded prior to calling
        /// this method!
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
                    // Attempt to convert the CMS color space signature
                    // to an XL Toolbox ColorSpaceSignature enum value.
                    // If this fails, cpvm will remain null.
                    c = cmsGetColorSpace(h);
                    ColorSpace colorSpace = c.ToColorSpace();
                    cpvm = new ColorProfileViewModel(fileName);
                    cpvm.ColorSpace = colorSpace;
                }
                catch (NotImplementedException)
                { 
                    // NO-OP
                }
                finally
                {
                    cmsCloseProfile(h);
                }
            }
            return cpvm;
        }

        /// <summary>
        /// Creates a new ColorProfileViewModel from a given name, but only
        /// if the color profile's color space is known by the view model.
        /// Directory and file extension will be appended automatically.
        /// Note that the lcms2.dll must have been loaded prior to calling
        /// this method!
        /// </summary>
        /// <param name="fileName">Full file name of the ICS color profile.</param>
        /// <returns>Instance of ColorProfileViewModel, or null if the color
        /// profile has an unknown color space.</returns>
        public static ColorProfileViewModel CreateFromName(string name)
        {
            return CreateFromFile(
                System.IO.Path.Combine(
                    Pinvoke.GetColorDirectory(),
                    name + Lcms.Constants.COLOR_PROFILE_EXTENSION
                )
            );
        }

        #endregion

        #region Public properties

        public ColorSpace ColorSpace { get; private set; }

        public string Name { get; private set; }

        #endregion

        #region Public methods

        /// <summary>
        /// Transforms the pixels of a FreeImageBitmap using a standard color profile
        /// for the bitmap's color space and the current color profile.
        /// </summary>
        /// <param name="freeImageBitmap">FreeImageBitmap whose pixels to convert.</param>
        /// <remarks>
        /// Since Excel does not support color management at all, conversions are done
        /// using standard color profiles that LittleCMS provides.
        /// </remarks>
        public void TransformFromStandardProfile(FreeImageAPI.FreeImageBitmap freeImageBitmap)
        {
            cmsHProfile standardProfile = CreateStandardProfile(freeImageBitmap);
            if (standardProfile == cmsHProfile.Zero)
            {
                throw new InvalidOperationException(
                    "Unable to create standard profile for " + freeImageBitmap.ColorType.ToString());
            }

            cmsHProfile targetProfile = cmsOpenProfileFromFile(GetPathFromName(), "r");
            if (targetProfile == cmsHProfile.Zero)
            {
                throw new InvalidOperationException(
                    "Unable to open desired color profile: " + Name);
            }

            // Create transform with perceptual intent (0) and no special options (0)
            cmsHTransform t = cmsCreateTransform(
                standardProfile, GetLcmsPixelFormat(freeImageBitmap),
                targetProfile, GetLcmsPixelFormat(ColorSpace),
                0, 0);
            if (t == cmsHTransform.Zero)
            {
                throw new InvalidOperationException("Unable to create CMS transform.");
            }
            UInt32 numPixels = (UInt32)freeImageBitmap.Size.Width * (UInt32)freeImageBitmap.Size.Height;
            cmsDoTransform(t, freeImageBitmap.Bits, freeImageBitmap.Bits, numPixels);
            freeImageBitmap.CreateICCProfile(GetIccBytes());

            cmsDeleteTransform(t);
            cmsCloseProfile(standardProfile);
            cmsCloseProfile(targetProfile);
        }

        #endregion

        #region Protected constructor

        /// <summary>
        /// Instantiates the view model given a color profile's file name.
        /// </summary>
        /// <param name="name">Complete file name to the color profile.</param>
        /// <exception cref="ArgumentException">if the color profile does not
        /// exist or is not accessible.</exception>
        /// <remarks>
        /// The constructor is protected in order to enforce the factory method,
        /// which takes care of restricting instantiation to profiles with color
        /// spaces that we can actually handle.
        /// </remarks>
        protected ColorProfileViewModel(string name)
            :base()
        {
            Name = SanitizeName(name);
        }

        #endregion

        #region Overrides

        public override string DisplayString
        {
            get
            {
                return Name;
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

        #region Private methods

        /// <summary>
        /// Removes path and extension from name.
        /// </summary>
        /// <returns>Sanitized name.</returns>
        private string SanitizeName(string name)
        {
            string s = System.IO.Path.GetFileName(name);
            if (s.ToLower().EndsWith(Lcms.Constants.COLOR_PROFILE_EXTENSION))
            {
                s = s.Remove(s.Length - Lcms.Constants.COLOR_PROFILE_EXTENSION.Length);
            }
            return s;
        }

        /// <summary>
        /// Returns a complete path for the color profile.
        /// </summary>
        /// <returns>Complete path to the color profile file.</returns>
        private string GetPathFromName()
        {
            return System.IO.Path.Combine(
                Pinvoke.GetColorDirectory(),
                Name + Lcms.Constants.COLOR_PROFILE_EXTENSION);
        }

        /// <summary>
        /// Creates a standard color profile for the color space of the given
        /// <paramref name="freeImageBitmap"/>.
        /// </summary>
        /// <param name="freeImageBitmap">FreeImageBitmap for whose color space
        /// to create a profile.</param>
        /// <returns>Handle to a standard profile, or zero if no standard profile
        /// could be created.</returns>
        private cmsHProfile CreateStandardProfile(FreeImageAPI.FreeImageBitmap freeImageBitmap)
        {
            switch (freeImageBitmap.ColorType)
            {
                case FreeImageAPI.FREE_IMAGE_COLOR_TYPE.FIC_RGB:
                    return cmsCreate_sRGBProfile();
                case FreeImageAPI.FREE_IMAGE_COLOR_TYPE.FIC_RGBALPHA:
                    return cmsCreate_sRGBProfile();
                default:
                    return cmsHProfile.Zero;
            }
        }

        /// <summary>
        /// Converts a ColorSpace value to an LCMS formatter constant, as defined
        /// in <see cref="XLToolbox.Export.Lcms.Formatters"/>.
        /// </summary>
        /// <param name="colorSpace">ColorSpace value to convert.</param>
        /// <returns>Resulting LCMS formatter constant.</returns>
        /// <exception cref="InvalidOperationException">if no predefined formatter
        /// constant exists for the given ColorSpace value.</exception>
        private cmsUInt32Number GetLcmsPixelFormat(ColorSpace colorSpace)
        {
            switch (colorSpace)
            {
                case Models.ColorSpace.Cmyk:
                    return Lcms.Formatters.TYPE_CMYK;
                case Models.ColorSpace.Rgb:
                    return Lcms.Formatters.TYPE_RGBA_8;
                case Models.ColorSpace.GrayScale:
                    return Lcms.Formatters.GRAY_8;
                default:
                    throw new InvalidOperationException(
                        "No LCMS formatter defined for " + colorSpace.ToString());
            }
        }

        /// <summary>
        /// Converts a FreeImage color type to an LCMS formatter constant, as defined
        /// in <see cref="XLToolbox.Export.Lcms.Formatters"/>.
        /// </summary>
        /// <param name="freeImageBitmap">FreeImageBitmap whose color type to convert.</param>
        /// <returns>Resulting LCMS formatter constant.</returns>
        /// <exception cref="InvalidOperationException">if no predefined formatter
        /// constant exists for the color type of the given FreeImageBitmap.</exception>
        private cmsUInt32Number GetLcmsPixelFormat(FreeImageAPI.FreeImageBitmap freeImageBitmap)
        {
            switch (freeImageBitmap.ColorType)
            {
                case FreeImageAPI.FREE_IMAGE_COLOR_TYPE.FIC_CMYK:
                    return Lcms.Formatters.TYPE_CMYK;
                case FreeImageAPI.FREE_IMAGE_COLOR_TYPE.FIC_RGBALPHA:
                    return Lcms.Formatters.TYPE_BGRA_8; // on Windows: BGRA, not RGBA!
                case FreeImageAPI.FREE_IMAGE_COLOR_TYPE.FIC_RGB:
                    return Lcms.Formatters.TYPE_BGR_8; // on Windows: BGR, not RGB!
                default:
                    throw new InvalidOperationException(
                        "No LCMS formatter defined for " + freeImageBitmap.ColorType.ToString());
            }
        }

        private byte[] GetIccBytes()
        {
            return System.IO.File.ReadAllBytes(GetPathFromName());
        }

        #endregion

        #region Private fields

        #endregion

        #region P/Invokes

        [DllImport("lcms2.dll", EntryPoint = "cmsOpenProfileFromFile")]
        private static extern cmsHProfile cmsOpenProfileFromFile(string iccProfile, string access);

        [DllImport("lcms2.dll", EntryPoint = "cmsCloseProfile")]
        private static extern cmsBool cmsCloseProfile(cmsHProfile hProfile);

        [DllImport("lcms2.dll", EntryPoint = "cmsGetColorSpace")]
        private static extern ColorSpaceSignature cmsGetColorSpace(cmsHProfile hProfile);

        [DllImport("lcms2.dll", EntryPoint = "cmsCreate_sRGBProfile")]
        private static extern cmsHProfile cmsCreate_sRGBProfile();

        [DllImport("lcms2.dll", EntryPoint = "cmsCreateTransform")]
        private static extern cmsHTransform cmsCreateTransform(
            cmsHProfile Input, cmsUInt32Number InputFormat,
            cmsHProfile Output, cmsUInt32Number OutputFormat,
            cmsUInt32Number Intent, cmsUInt32Number dwFlags);

        [DllImport("lcms2.dll", EntryPoint = "cmsDeleteTransform")]
        private static extern void cmsDeleteTransform(cmsHTransform hTransform);

        [DllImport("lcms2.dll", EntryPoint = "cmsDoTransform")]
        private static extern void cmsDoTransform(
            cmsHTransform hTransform,
            IntPtr InputBuffer,
            IntPtr OutputBuffer,
            cmsUInt32Number	Size);

        #endregion
    }
}
