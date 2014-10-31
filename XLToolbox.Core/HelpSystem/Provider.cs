using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.HelpSystem
{
    public enum Topic
    {
        WhatsNew,
        Donate
    };

    public class Provider
    {
        /// <summary>
        /// Invokes the default browser and navigates to a help topic.
        /// </summary>
        /// <param name="topic">Help topic to show</param>
        public static void Show(Topic topic)
        {
            string url = (string)Properties.HelpTopics.Default[topic.ToString()];
            System.Diagnostics.Process.Start(Properties.Settings.Default.HelpUrl + url);
        }
    }
}
