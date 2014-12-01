using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm;
using System.Windows.Resources;
using System.Windows;
using Bovender.Mvvm.ViewModels;
using System.IO;

namespace Bovender.HtmlFiles
{
    /// <summary>
    /// View model for a HTML file.
    /// </summary>
    public class HtmlFileViewModel : ViewModelBase
    {
        #region Constructors

        /// <summary>
        /// Constructor that loads a HTML file from a qualified pack URI.
        /// </summary>
        /// <param name="packUri">Valid absolute pack URI
        /// (e.g. "pack://application:,,,/ASSEMBLYNAME;component/FILENAME"</param>
        /// <remarks>The build action of the file to be loaded must be "Resource".</remarks>
        public HtmlFileViewModel(string packUri)
        {
            _packUri = packUri;
            HtmlStream = Application.GetResourceStream(new Uri(packUri)).Stream;
        }

        /// <summary>
        /// Constructor that loads a HTML file given an assembly name
        /// and a path to a file that has its build action set to "Resource".
        /// </summary>
        /// <param name="assemblyName">Assembly name</param>
        /// <param name="filePath">HTML file in the assembly (build action must be "Resource")</param>
        public HtmlFileViewModel(string assemblyName, string filePath)
            : this(String.Format(
                    "pack://application:,,,/{0};component/{1}",
                    assemblyName, filePath
                ))
        { }

        /// <summary>
        /// Loads a HTML file given an assembly name and file name and sets a caption.
        /// </summary>
        /// <param name="caption">Custom caption</param>
        /// <param name="assemblyName">Assembly name</param>
        /// <param name="filePath">HTML file in the assembly (build action must be "Resource")</param>
        public HtmlFileViewModel(string caption, string assemblyName, string filePath)
            : this(assemblyName, filePath)
        {
            Caption = caption;
        }

        #endregion

        #region Properties

        public string Caption { get; set; }
        public Stream HtmlStream { get; set; }

        #endregion

        #region Private fields

        readonly string _packUri;

        #endregion

        #region Implementation of ViewModelBase's abstract methods

        public override bool IsViewModelOf(object model)
        {
            return String.Equals(model, _packUri);
        }

        #endregion
    }
}
