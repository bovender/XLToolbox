using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Help
{
    public enum Topic
    {
        WhatsNew,
        Donate
    };

    public class Provider
    {
#if !DEBUG
        private const string _baseUrl = "http://xltoolbox.sf.net/";
#else
        private const string _baseUrl = "http://xltb.vhost/";
#endif

        /// <summary>
        /// Invokes the default browser and navigates to a help topic.
        /// </summary>
        /// <param name="topic">Help topic to show</param>
        public static void Show(Topic topic)
        {
            string url = (string)Properties.Topics.Default[topic.ToString()];
            System.Diagnostics.Process.Start(_baseUrl + url);
        }
    }
}
