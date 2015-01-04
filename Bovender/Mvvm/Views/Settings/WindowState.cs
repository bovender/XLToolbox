/* WindowState.cs
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
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media;

namespace Bovender.Mvvm.Views.Settings
{
    /// <summary>
    /// Provides an attached property to enable Windows to save
    /// and restore their screen position and state.
    /// </summary>
    public static class WindowState
    {
        #region Attached property "Save"

        public static readonly DependencyProperty SaveProperty = DependencyProperty.RegisterAttached(
            "Save", typeof(bool), typeof(WindowState), new PropertyMetadata(OnSavePropertyChanged));

        [AttachedPropertyBrowsableForType(typeof(Window))]
        public static bool GetSave(UIElement element)
        {
            return (bool)element.GetValue(SaveProperty);
        }

        public static void SetSave(UIElement element, bool value)
        {
            element.SetValue(SaveProperty, value);
        }

        private static void OnSavePropertyChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            if (e.Property == SaveProperty)
            {
                Window window = obj as Window;
                if (window != null)
                {
                    if ((bool)e.NewValue == true)
                    {
                        window.Closed += SaveWindowGeometry;
                        window.Initialized += LoadWindowGeometry;
                    }
                    else
                    {
                        window.Closed -= SaveWindowGeometry;
                        window.Initialized -= LoadWindowGeometry;
                    }
                }
            }
        }

        #endregion

        #region Attached property "CenterScreen"

        public static readonly DependencyProperty CenterScreenProperty = DependencyProperty.RegisterAttached(
            "CenterScreen", typeof(bool), typeof(WindowState), new PropertyMetadata(OnCenterScreenPropertyChanged));

        /// <summary>
        /// Centers the Window that the property is attached to on the screen that the
        /// window is currently being displayed on. This also works with WPF windows that
        /// are creaed by a VSTO addin where the native "WindowStartupLocation" does not
        /// work.
        /// </summary>
        /// <param name="element">UI element to center on screen, must be a Window</param>
        /// <returns>Attached property "CenterScreen"</returns>
        [AttachedPropertyBrowsableForType(typeof(Window))]
        public static bool GetCenterScreen(UIElement element)
        {
            return (bool)element.GetValue(CenterScreenProperty);
        }

        public static void SetCenterScreen(UIElement element, bool value)
        {
            element.SetValue(CenterScreenProperty, value);
        }

        private static void OnCenterScreenPropertyChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            if (e.Property == CenterScreenProperty)
            {
                Window window = obj as Window;
                if (window != null)
                {
                    if ((bool)e.NewValue == true)
                    {
                        // Must not use the Initialized event here
                        // Only when Loaded is raised, the window handle will be accessible
                        window.Loaded += CenterWindow;
                    }
                }
            }
        }

        private static void CenterWindow(object sender, EventArgs e)
        {
            Window w = sender as Window;
            Screen screen = Screen.FromHandle(new WindowInteropHelper(w).Handle);
            PresentationSource ps = PresentationSource.FromVisual(w);
            CompositionTarget ct = ps.CompositionTarget;
            Matrix t = ct.TransformFromDevice;
            Point screenGeo = t.Transform(new Point(screen.Bounds.Width, screen.Bounds.Height));
            double screenWidth = screenGeo.X;
            double screenHeight = screenGeo.Y;
            w.Left = Math.Max((screenWidth - w.ActualWidth) / 2, 0);
            w.Top = Math.Max((screenHeight - w.ActualHeight) / 2, 0);
        }

        #endregion

        #region Event handlers

        static void LoadWindowGeometry(object sender, EventArgs e)
        {
            WindowSettings s = WindowSettingsFactory(sender);
            Window w = sender as Window;
            s.Reload();
            if (s.Rect != Rect.Empty)
            {
                w.Left = s.Rect.Left;
                w.Top = s.Rect.Top;
                w.Width = s.Rect.Width;
                w.Height = s.Rect.Height;
                SanitizeWindowGeometry(w);

                // Since the CenterScreen function is hooked to the
                // Window.Loaded event, which is raised after the
                // Window.Initialize event that leads to the current
                // function call, we need to make sure that the window
                // is not centered if we use previously saved geometry.
                w.Loaded -= CenterWindow;
            }
            w.WindowState = s.State;
        }

        static void SaveWindowGeometry(object sender, EventArgs e)
        {
            WindowSettings s = WindowSettingsFactory(sender);
            Window w = sender as Window;
            s.Rect = new Rect(w.Left, w.Top, w.Width, w.Height);
            s.State = w.WindowState;
            s.Screen = Screen.FromHandle(new WindowInteropHelper(w).Handle);
            s.Save();
        }

        #endregion

        #region Private methods

        static WindowSettings WindowSettingsFactory(object obj)
        {
            Window w = obj as Window;
            if (w != null)
            {
                return new WindowSettings(w);
            }
            else
            {
                throw new InvalidOperationException("The WindowState.Save property must be attached to Windows only.");
            }
        }

        /// <summary>
        /// Adjusts a window's geometry for the current screen resolution.
        /// </summary>
        /// <param name="window"></param>
        static void SanitizeWindowGeometry(Window window)
        {
            Window w = window;
            double vleft = SystemParameters.VirtualScreenLeft;
            double vright = vleft + SystemParameters.VirtualScreenWidth;
            double vtop = SystemParameters.VirtualScreenTop;
            double vbottom = vtop + SystemParameters.VirtualScreenHeight;

            // Make sure that the top left corner of the window
            // is inside the virtual screen.
            // Note: Cannot use Window.ActualWidth and Window.ActualHeight
            // as these won't be initialized yet when this function is called
            // as a consequence of the Window.Initialize event.
            if (w.Left < vleft)
            {
                w.Left = vleft;
            }
            if (w.Left > vright - w.Width)
            {
                w.Left = vright - w.Width;
            }
            if (w.Top < vtop)
            {
                w.Top = vtop;
            }
            if (w.Top > vbottom - w.Height)
            {
                w.Top = vbottom - w.Height;
            }

            // TODO: Make sure the window location is mapped to an actual screen
            // The virtual screen represents a bounding rectangle that does not
            // necessarily need to be entirely mapped to physical screens, e.g.
            // if two monitors are aligned diagonally.
        }

        #endregion
           
    }
}
