using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bovender
{
    /// <summary>
    /// Provides easy public access to this assembly's settings file,
    /// e.g. the user's e-mail address used for submitting exception
    /// reports.
    /// </summary>
    /// <remarks>
    /// Changed properties need to be explicitly made persistant by
    /// calling <see cref="Save()"/>.
    /// </remarks>
    public static class Settings
    {
        public static string User
        {
            get { return Properties.Settings.Default.User; }
            set
            {
                Properties.Settings.Default.User = value;
            }
        }

        public static string Email
        {
            get { return Properties.Settings.Default.Email; }
            set
            {
                Properties.Settings.Default.Email = value;
            }
        }

        public static bool CcUser
        {
            get { return Properties.Settings.Default.CcUser; }
            set
            {
                Properties.Settings.Default.CcUser = value;
            }
        }

        public static void Save()
        {
            Properties.Settings.Default.Save();
        }
    }
}
