using System;
using System.Collections.Specialized;
using System.Reflection;
using XLToolbox.Versioning;

namespace XLToolbox.ExceptionHandler
{
    public class ExceptionViewModel : Bovender.ExceptionHandler.ExceptionViewModel
    {
        #region Additional properties for the exception report

        public string ExcelVersion
        {
            get
            {
                return Excel.Instance.ExcelInstance.Application.Version;
            }
        }

        public string ToolboxVersion
        {
            get
            {
                return SemanticVersion.CurrentVersion().ToString();
            }
        }

        public string FreeImageVersion
        {
            get
            {
                if (FreeImageAPI.FreeImage.IsAvailable())
                {
                    return FreeImageAPI.FreeImage.GetVersion();
                }
                else
                {
                    return "Not available";
                }
            }
        }

        #endregion

        #region constructor

        public ExceptionViewModel(Exception e) : base(e) { }

        #endregion

        #region Overrides

        protected override NameValueCollection GetPostValues()
        {
            NameValueCollection v = base.GetPostValues();
            v["excel_version"] = ExcelVersion;
            v["excel_bitness"] = ProcessBitness;
            v["freeimage_version"] = FreeImageVersion;
            v["toolbox_version"] = ToolboxVersion;
            return v;
        }

        protected override Uri GetPostUri()
        {
            return new Uri(Properties.Settings.Default.ExceptionPostUrl);
        }

        #endregion
    }
}
