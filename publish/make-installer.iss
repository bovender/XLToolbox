; #define DEBUG
#define VERSIONFILE "..\VERSION"
#define PUB_YEARS "2016"
#define TARGET_HOST "excel"
#define APP_GUID "{{35AD3250-5F75-4C7D-BCE0-41377E280430}"
#define ADDIN_NAME "Daniel's XL Toolbox"
#define ADDIN_SHORT_NAME "XLToolbox"
#define COMPANY "Daniel Kraus (bovender)"
#define DESCRIPTION "Scientific add-in for Microsoft Excel."
#define HOMEPAGE "https://www.xltoolbox.net"
#define HOMEPAGE_SUPPORT "https://www.xltoolbox.net/support.html"
#define HOMEPAGE_UPDATES "https://github.com/bovender/XLToolbox/releases"
#define VSTOFILE "XLToolboxForExcel.vsto"
#define OUTPUTDIR "release\"
#define LOGFILE "XLTOOLBOX-INSTALL-LOG.TXT"
#define REGKEY "XL Toolbox NG"

#define SETUPFILESDIR "setup-files\"
#define LICENSE_FILE "license.rtf"
#define INSTALLER_ICO "xltoolbox.ico"
#define INSTALLER_IMAGE_LARGE "installbanner.bmp"
#define INSTALLER_IMAGE_SMALL "icon-48.bmp"

#define DEVENV "C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\devenv.com"
#define CSPROJ "x:\Code\xltoolbox\NG\NG.sln"

#ifdef DEBUG
  #define SOURCEDIR "..\XLToolboxForExcel\bin\Debug"
  #expr Exec(DEVENV, CSPROJ + " /Build Debug")
#else
  #define SOURCEDIR "..\XLToolboxForExcel\bin\Release"
  #expr Exec(DEVENV, CSPROJ + " /Build Release")
#endif

#include "VstoAddinInstaller\vsto-installer.iss"

[Run]
Filename: "{code:LegacyUninstallerPath}"; Parameters: "/SILENT"; Description: "Remove legacy version (required)"; StatusMsg: "Uninstalling legacy version..."; Check: IsLegacyInstalled; Flags: runminimized;
Filename: "setx"; Parameters: "VSTO_SUPPRESSDISPLAYALERTS 0"; Description: "Enable add-in loading error messages"; StatusMsg: "Enabling error messages while add-in is being loaded..."; Flags: runhidden nowait;
