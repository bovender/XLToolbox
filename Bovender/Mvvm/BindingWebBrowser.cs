using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Resources;
using System.IO;

namespace Bovender.Mvvm
{
    /// <summary>
    /// Provides attached properties to facilitate data binding of a
    /// WebBrowser control.
    /// </summary>
    /// <remarks>
    /// Usage:
    ///     <![CDATA[
    ///     * <WebBrowser b:BindingWebBrowser.Html="{Binding HtmlString}" />
    ///     * <WebBrowser b:BindingWebBrowser.Stream="{Binding Htmlstream}" />
    ///     ]]>
    /// </remarks>
    public static class BindingWebBrowser
    {
        #region Attached property 'Html'

        public static readonly DependencyProperty HtmlProperty =
            DependencyProperty.RegisterAttached(
                "Html",
                typeof(string),
                typeof(BindingWebBrowser),
                new UIPropertyMetadata(null, HtmlPropertyChanged));

        public static string GetHtml(DependencyObject obj)
        {
            return (string)obj.GetValue(HtmlProperty);
        }

        public static void SetHtml(DependencyObject obj, string value)
        {
            obj.SetValue(HtmlProperty, value);
        }

        public static void HtmlPropertyChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            WebBrowser browser = o as WebBrowser;
            if (browser != null)
            {
                browser.NavigateToString(e.NewValue as string);
            }
        }

        #endregion

        #region Attached property 'Stream'

        public static readonly DependencyProperty StreamProperty =
            DependencyProperty.RegisterAttached(
                "Stream",
                typeof(Stream),
                typeof(BindingWebBrowser),
                new UIPropertyMetadata(null, StreamPropertyChanged));

        public static string GetStream(DependencyObject obj)
        {
            return (string)obj.GetValue(StreamProperty);
        }

        public static void SetStream(DependencyObject obj, string value)
        {
            obj.SetValue(StreamProperty, value);
        }

        public static void StreamPropertyChanged(DependencyObject o, DependencyPropertyChangedEventArgs e)
        {
            WebBrowser browser = o as WebBrowser;
            if (browser != null)
            {
                Stream s = e.NewValue as Stream;
                s.Seek(0, SeekOrigin.Begin);
                browser.NavigateToStream(s);
            }
        }

        #endregion
    }
}
