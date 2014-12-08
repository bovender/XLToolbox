using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Resources;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace XLToolbox
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("XLToolbox.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Button_OnAction(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "ButtonAbout": Dispatcher.Execute(Command.About);
                    break;
                case "ButtonCheckForUpdate": Dispatcher.Execute(Command.CheckForUpdates);
                    break;
                case "ButtonTestError": Dispatcher.Execute(Command.ThrowError);
                    break;
                case "ButtonSheetList": Dispatcher.Execute(Command.SheetList);
                    break;
                case "ButtonExportSelection": Dispatcher.Execute(Command.ExportSelection);
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// Returns an Image object for the ribbon.
        /// </summary>
        /// <remarks>
        /// The image file is expected to be a WPF resource file, not an embedded resource.
        /// To be consistent accross the application which uses WPF resources for its WPF
        /// windows, all images are to be built as resources rather than embedded resources.
        /// </remarks>
        /// <param name="imageId">The file name (without path) of the image.</param>
        /// <returns>Image object</returns>
        public object Ribbon_LoadImage(string imageId)
        {
            string initPackScheme = System.IO.Packaging.PackUriHelper.UriSchemePack;
            StreamResourceInfo sri = Application.GetResourceStream(
                new Uri(@"pack://application:,,,/XLToolbox;component/Resources/images/" + imageId));
            return Image.FromStream(sri.Stream);
        }

        public string Control_GetLabel(Office.IRibbonControl control)
        {
            return LookupResourceString(control.Id + "Label");
        }

        public string Control_GetSupertip(Office.IRibbonControl control)
        {
            return LookupResourceString(control.Id + "Supertip");
        }

        private string LookupResourceString(string name)
        {
            return RibbonStrings.ResourceManager.GetString(name);
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public bool Group_IsVisibleInDebugOnly(Office.IRibbonControl control)
        {
#if DEBUG
            return true;
#else
            return false;
#endif
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
