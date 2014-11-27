using System.Configuration;
using System.Windows;

namespace Bovender.Mvvm.Views.Settings
{
    public class WindowSettings : ApplicationSettingsBase
    {
        #region Constructor

        public WindowSettings(Window window) : base(window.ToString()) { }

        #endregion

        #region Properties

        /// <summary>
        /// The rectangle describing the window's coordinates.
        /// </summary>
        [UserScopedSetting]
        public Rect Rect
        {
            get
            {
                // Cannot return this["Rect"] as Rect
                // because Rect is not nullable
                object obj = this["Rect"];
                if (obj == null) {
                    return Rect.Empty;
                }
                else {
                    return (Rect)obj;
                }
            }
            set
            {
                this["Rect"] = value;
            }
        }

        [UserScopedSetting]
        public System.Windows.WindowState State
        {
            get
            {
                object obj = this["State"];
                if (obj == null)
                {
                    return System.Windows.WindowState.Normal;
                }
                else
                {
                    return (System.Windows.WindowState)obj;
                }
            }
            set
            {
                this["State"] = value;
            }
        }

        [UserScopedSetting]
        public System.Windows.Forms.Screen Screen
        {
            get
            {
                return (System.Windows.Forms.Screen)this["Screen"];
            }
            set
            {
                this["Screen"] = value;
            }
        }

        #endregion

    }
}
