; Inno Setup script for Daniel's XL Toolbox
; (c) 2008-2015 Daniel Kraus
; Apache License Version 2.0

[Setup]
; #define DEBUG

; Read the semantic and the installer file version from the VERSION file
#define FILE_HANDLE FileOpen("..\VERSION")
#define SEMVER FileRead(FILE_HANDLE)
#define VER FileRead(FILE_HANDLE)
#expr FileClose(FILE_HANDLE)

#define YEAR "2015"
#define DEV "Daniel Kraus"
#define LOGFILE "INST-LOG.TXT"
#define REGKEY "Software\Microsoft\Office\Excel\Addins\XL Toolbox NG"
#define APPNAME "Daniel's XL Toolbox NG"
#define SLOGAN "Scientific add-in for Microsoft Excel."
#define UNINSTALLDIR "{app}\setup"
#define ADDINNAME "XLToolboxForExcel"
#define DOTNETSHA1 "58da3d74db353aad03588cbb5cea8234166d8b99"
#define VSTORSHA1 "ad1dcc5325cb31754105c8c783995649e2208571"

#ifndef DEBUG
  #define SOURCEDIR "Release"
	#define VSTORURL "http://download.microsoft.com/download/2/E/9/2E9D2603-6D1F-4B12-BD37-DB1410B23597/vstor_redist.exe"
	#define DOTNETURL "http://download.microsoft.com/download/9/5/A/95A9616B-7A37-4AF6-BC36-D6EA96C8DAAE/dotNetFx40_Full_x86_x64.exe"
	; Build the solution with Release configuration
	#expr Exec("C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\devenv.com", \
				"x:\Code\xltoolbox\NG\NG.sln /Build Release")
#else
  #define SOURCEDIR "Debug"
	#define VSTORURL "http://vhost/vstor_redist.exe"
	#define DOTNETURL "http://vhost/dotNetFx40_Full_x86_x64.exe"
#endif

; Specific AppID - NEVER CHANGE THIS!
AppId={{35AD3250-5F75-4C7D-BCE0-41377E280430}
                      
; Compiler info
OutputDir=release
#ifndef DEBUG
	OutputBaseFilename=XL_Toolbox_{#SEMVER}
#else
	OutputBaseFilename=XL_Toolbox_debug
#endif
Compression=lzma
InternalCompressLevel=max
SolidCompression=true
LicenseFile=setup-files\license.rtf

; Application name, version, etc.
AppName={#APPNAME}
AppVersion={#SEMVER}
AppPublisher={#DEV}
AppCopyright=2008-{#YEAR} {#DEV}
AppComments={#SLOGAN}
VersionInfoDescription={#SLOGAN}
VersionInfoVersion={#VER}
VersionInfoProductVersion={#VER}
VersionInfoTextVersion={#SEMVER}
VersionInfoProductTextVersion={#SEMVER}
AppContact={#DEV}
AppPublisherURL=http://www.xltoolbox.net
AppSupportURL=http://www.xltoolbox.net/support.html
AppUpdatesURL=http://sf.net/projects/xltoolbox/files/latest

; Setup wizard
SetupIconFile=setup-files\xltoolbox.ico
WizardImageFile=setup-files\installbanner.bmp
WizardSmallImageFile=setup-files\icon-48.bmp
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64
DisableProgramGroupPage=true
DisableDirPage=false
DisableReadyPage=false
LanguageDetectionMethod=locale
TimeStampsInUTC=false
PrivilegesRequired=lowest
WizardImageStretch=false
WizardImageBackColor=clWhite
; ChangesAssociations=true
AppendDefaultDirName=false
DefaultDirName={code:SuggestInstallDir}
CreateAppDir=true
CloseApplicationsFilter=*.*
DefaultDialogFontName=Segoe UI

; Uninstall
UninstallDisplayIcon={#UNINSTALLDIR}\xltoolbox.ico
UninstallFilesDir={#UNINSTALLDIR}
	
; Inno Downloader Plugin is required for this
; Note that this include directive MUST be located at the end of the [setup]
; section.
#include <idp.iss>

[Languages]
Name: en; MessagesFile: compiler:Default.isl; 
Name: de; MessagesFile: compiler:Languages\German.isl; 
; Name: Brasileiro; MessagesFile: compiler:Languages\BrazilianPortuguese.isl;
; Name: Portugues; MessagesFile: compiler:Languages\Portuguese.isl;
#include "c:\Program Files (x86)\Inno Download Plugin\Unicode\idplang\german.iss"

[Files]
Source: "..\XLToolboxForExcel\bin\{#SOURCEDIR}\*"; DestDir: "{app}"; Flags: ignoreversion createallsubdirs recursesubdirs
; Source: "setup-files\source.zip"; DestDir: "{app}"
Source: "setup-files\xltoolbox.ico"; DestDir: "{#UNINSTALLDIR}"
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Registry]
; Keys for single-user install (HKCU)
Check: not IsMultiUserInstall; ValueName: Description; ValueData: {#SLOGAN}; ValueType: string; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: FriendlyName; ValueData: {#APPNAME}; ValueType: string; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: LoadBehavior; ValueData: 3; ValueType: dword; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: Warmup; ValueData: 1; ValueType: dword; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: Manifest; ValueData: file:///{code:ConvertSlash|{app}}/{#ADDINNAME}.vsto|vstolocal; ValueType: string; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey

; Same keys again, this time for multi-user install (HKLM)
Check: IsMultiUserInstall; ValueName: Description; ValueData: {#SLOGAN}; ValueType: string; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: FriendlyName; ValueData: {#APPNAME}; ValueType: string; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: LoadBehavior; ValueData: 3; ValueType: dword; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: Warmup; ValueData: 1; ValueType: dword; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: Manifest; ValueData: file:///{code:ConvertSlash|{app}}/{#ADDINNAME}.vsto|vstolocal; ValueType: string; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey

[Run]
; Filename: http://xltoolbox.sourceforge.net/welcome.html; Flags: shellexec nowait; Check: ShowWelcomePageInBrowser

[UninstallRun]
; Filename: http://xltoolbox.sourceforge.net/uninstall.html; Flags: shellexec nowait

[CustomMessages]
; General messages [EN]
en.DevVer=Developmental version
en.DevVerSubcaption=Please acknowledge that this is development in progress.
en.DevVerDesc=Please acknowledge that this is development in progress. Please note that this is a developmental version of the XL Toolbox. Expect things to work differently from what you expect to not work at all, to crash your system, and to change in any subsequent version.
en.DevVerCheckbox=I understand and expect things to not work and to crash at any time.
en.DevVerMsgBox=You must check the box to acknowledge that this is a development version in order to proceed.
en.SingleOrMulti=Single-user or system-wide install
en.SingleOrMultiSubcaption=Install for the current user only or for all users
en.SingleOrMultiDesc=Please indicate the scope of this installation:
en.SingleOrMultiSingle=Single user (only for me)
en.SingleOrMultiAll=All users (system-wide)
en.Excel2007Required=Daniel's XL Toolbox NG requires Excel 2007 or later. Please download and install the legacy version (e.g. 6.52) of this add-in which also works with Excel 2003. Setup will now terminate.

; CannotInstallPage [EN]
en.CannotInstallCaption=Administrator privileges required
en.CannotInstallDesc=You do not have the necessary rights to install additional required runtime files.
en.CannotInstallMsg=Additional runtime files from Microsoft are required to run the XL Toolbox. You may continue the installation, but the XL Toolbox won't start unless the required runtime files are installed by an administrator. Note: On Windows Vista and newer, right-click the installer file and choose 'Run as administrator'.
en.CannotInstallCont=Continue anyway, although it won't work without the required runtime files
en.CannotInstallAbort=Abort the installation (come back when the admin has installed the files)

; DownloadInfoPage [EN]
en.RequiredCaption=Additional runtime files required
en.RequiredDesc=Additional runtime files for the .NET framework from Microsoft are required in order to run the XL Toolbox.
en.RequiredMsg=%d file(s) totalling about %s MiB need to be downloaded from the Microsoft servers. Click 'Next' to start downloading.

; InstallInfoPage [EN]
en.InstallCaption=Runtime files downloaded
en.InstallDesc=The required runtime files are ready to install.
en.InstallMsg=Click 'Next' to beginn the installation.

en.StillNotInstalled=The required additional runtime files are still not installed. Setup cannot continue. You may try again, or abort the XL Toolbox installation.
en.DownloadNotValidated=A downloaded file has unexpected content. It may have not been downloaded correctly, or someone might have hampered with it. You may click 'Back' and then 'Next' to download it again.

; General messages [DE]
de.DevVer=Entwicklerversion
de.DevVerSubcaption=Bestätigen Sie, daß Sie die Entwicklerversion installieren wollen.
de.DevVerDesc=Beachten Sie bitte, daß es sich hierbei um eine Entwicklerversion handelt. Das Add-in kann sich anders verhalten, als Sie es erwarten, kann vielleicht Ihren Computer zum Absturz bringen, und kann in der nächsten Version ganz anders sein.
de.DevVerCheckbox=Ich verstehe, daß es sich um eine instabile Entwicklerversion handelt.
de.DevVerMsgBox=Sie müssen bestätigen, daß Sie die Entwicklerversion installieren wollen.
de.SingleOrMulti=Einzelner oder alle Benutzer
de.SingleOrMultiSubcaption=Geben Sie an, für wen die Installation sein soll
de.SingleOrMultiDesc=Bitte geben Sie an, ob die Toolbox nur für Sie oder für alle Benutzer installiert werden soll.
de.SingleOrMultiSingle=Ein Benutzer (nur für mich)
de.SingleOrMultiAll=Alle Benutzer (systemweit)
de.Excel2007Required=Daniel's XL Toolbox NG erfordert mindestens Excel 2007. Um das Add-in auch mit Excel 2003 zu verwenden, laden Sie bitte die alte, stabile Version 6.52 herunter.

; "Download required" messages (.NET and VSTOR runtimes) [DE]
de.CannotInstallCaption=Administratorrechte benötigt
de.CannotInstallDesc=Sie haben nicht die erforderlichen Benutzerrechte, um weitere benötigte Laufzeitdateien zu installieren.
de.CannotInstallMsg=Sie können mit der Installation fortfahren, aber die Toolbox wird nicht starten, solange die VSTO-Laufzeitdateien nicht von einem Admin installiert wurden. Tipp: Wenn Sie Windows Vista oder neuer verwenden, klicken Sie mit der rechten Maustaste auf die Installationsdatei und wählen "Als Administrator ausführen".
de.CannotInstallCont=Trotzdem installieren, obwohl es nicht funktionieren wird
de.CannotInstallAbort=Installation abbrechen

; DownloadInfoPage [EN]
de.RequiredCaption=Weitere Laufzeitdateien erforderlich
de.RequiredDesc=Weitere Laufzeitdateien für das .NET-Framework von Microsoft werden benötigt, um das Toolbox-Addin verwenden zu können.
de.RequiredMsg=%d Datei(en) mit ca. %s MiB muß/müssen von den Microsoft-Servern heruntergeladen werden. Klicken Sie 'Weiter', um den Download zu beginnen.

; InstallInfoPage [EN]
de.InstallCaption=Weitere .NET-Laufzeitdateien heruntergeladen
de.InstallDesc=Die zusätzlichen benötigten Dateien von Microsoft können jetzt installiert werden.
de.InstallMsg=Klicken Sie 'Weiter', um mit der Installation zu beginnen.

de.StillNotInstalled=Die zusätzlichen benötigten Dateien wurden leider nicht korrekt installiert, so daß die Toolbox-Installation nicht fortgesetzt werden kann.
de.DownloadNotValidated=Es wurde unerwarteter Inhalt in einer heruntergeladenen Datei gefunden. Die Installation kann so nicht fortgesetzt werden. Sie können aber 'Zurück' und dann 'Weiter' klicken, um den Download neu zu beginnen.

[Code]
const
	maxExcel = 24; //< highest Excel version number to check for.
	
var
	PageDevelopmentInfo: TInputOptionWizardPage;
	PageSingleOrMultiUser: TInputOptionWizardPage;
	PageCannotInstall: TInputOptionWizardPage;
	PageDownloadInfo: TOutputMsgWizardPage;
	PageInstallInfo: TOutputMsgWizardPage;
	prerequisitesChecked: boolean;
	prerequisitesMet: boolean;

  
/// Returns true if running on a zero client. The algorithm has only been
/// tested for VMware Horizon/Teradici clients.
function IsZeroClient(): boolean;
var
  protocol: string;
begin
  if RegQueryStringValue(HKEY_CURRENT_USER, 'Volatile Environment',
    'ViewClient_Protocol', protocol) then
  begin
    result := Uppercase(protocol) = 'PCOIP';
  end;
end;

/// Returns true if the target directory chooser should be shown or
/// not: This is the case if running on a zero client, or if the
/// current user is an administrator.
function ShouldShowDirPage(): boolean;
begin
  result := IsAdminLoggedOn or IsZeroClient;
end;
  
/// Returns the path for the Wow6432Node registry tree if the current operating
/// system is 64-bit, i.e., simulates WOW64 redirection.
function GetWowNode(): string;
begin
	if IsWin64 then
	begin
		result := 'Wow6432Node\';
	end
	else
	begin
		result := '';
	end;
end;
	
/// Checks if a given Excel version is installed
function IsExcelVersionInstalled(version: integer): boolean;
var key: string;
var lookup1, lookup2: boolean;
begin
	key := 'Microsoft\Office\' + IntToStr(version) + '.0\Excel\InstallRoot';
	lookup1 := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\' + GetWowNode + key);
	
	// If checking for version >= 14.0 ("2010"), which was the first version
	// that was produced in both 32-bit and 64-bit, on a 64-bit system we
	// also need to check a path without  'Wow6434Node'.
	if IsWin64 and (version >= 14) then
	begin
		lookup2 := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\' + key);
	end;
	
	result := lookup1 or lookup2;
end;

/// Checks if only Excel 2007 is installed
function IsOnlyExcel2007Installed(): boolean;
var
	i: integer;
begin
	result := IsExcelVersionInstalled(12);
	
	// Iterate through all
	for i := 14 to maxExcel do
	begin
		if IsExcelVersionInstalled(i) then
		begin
			result := false;
			break;
		end;
	end;
end;

/// Checks if hotfix KB976477 is installed. This hotfix
/// is required to make Excel 2007 recognize add-ins in
/// the HKLM hive as well.
function IsHotfixInstalled(): boolean;
begin
  result := RegKeyExists(HKEY_LOCAL_MACHINE,
		'SOFTWARE\Microsoft\Windows\Current Version\Uninstall\KB976477');
end;

/// Retrieves the build number of an installed Office version
/// in OutBuild. Returns true if the requested Office version
/// is installed and false if it is not installed.
function GetOfficeBuild(OfficeVersion: integer; var OutBuild: integer): boolean;
var
	key: string;
	value: string;
	build: string;
begin
	key := 'SOFTWARE\' + GetWowNode + 'Microsoft\Office\' +
					IntToStr(OfficeVersion) + '.0\Common\ProductVersion';
	if RegQueryStringValue(HKEY_LOCAL_MACHINE, key, 'LastProduct', value) then
	begin
		// Office build numbers always have 4 digits, at least as of Feb. 2015;
		// from a string '14.0.1234.5000' simply copy 4 characters from the 5th
		// position to get the build number. TODO: Make this future-proof.
		build := Copy(value, 6, 4);
		Log('Found ProductVersion "' + value + '" for queried Office version '
			+ IntToStr(OfficeVersion) + ', extracted build number ' + build);
		OutBuild := StrToInt(build);
		result := true;
	end
	else
		Log('Did not find LastProduct key for Office version ' +
				IntToStr(OfficeVersion) + '.0.');
end;

/// Asserts if Office 2007 is installed. Does not check whether other Office
/// versions are concurrently installed.
function IsOffice2007Installed(): boolean;
begin
	result := IsExcelVersionInstalled(12);
	if result then Log('Detected Office 2007.');
end;

/// Asserts if Office 2010 without service pack is installed.
/// For build number, see http://support.microsoft.com/kb/2121559/en-us
function IsOffice2010NoSpInstalled(): boolean;
var
	build: integer;
begin
	if GetOfficeBuild(14, build) then
	begin
		result := build = 4763; // 4763 is the original Office 2007 build
		if result then
			Log('Detected Office 2010 without service pack (v. 14.0, build 4763)')
		else
			Log('Detected Office 2010, apparently with some service pack (build ' +
					IntToStr(build) + ').');
	end;
end;

/// Checks if the VSTO runtime is installed. This is relevant if only
/// Excel 2007 is installed. Since Office 2010, the CLR is
/// automatically included.
/// The presence of the VSTO runtime is indicated by the presence one of
/// four possible registry keys.
/// See: http://xltoolbox.sf.net/blog/2015/01/net-vsto-add-ins-getting-prerequisites-right
/// HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R (32-bit)
/// HKLM\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R (64-bit)
function IsVstorInstalled(): boolean;
var
	software, vstorPath: string;
begin
	software := 'SOFTWARE\';
	vstorPath := 'Microsoft\VSTO Runtime Setup\v4R';
	result := RegKeyExists(HKEY_LOCAL_MACHINE, software + GetWowNode + vstorPath);
end;

/// Checks if the .NET 4.0 (or 4.5) runtime is installed.
/// See https://msdn.microsoft.com/en-us/library/hh925568
function IsNetInstalled(): boolean;
begin
	result := RegKeyExists(HKEY_LOCAL_MACHINE, 
		'SOFTWARE\' + GetWowNode + 'Microsoft\NET Framework Setup\NDP\v4');
end;

/// Asserts if the VSTO runtime for .NET 4.0 redistributable needs to be
/// downloaded and installed.
/// If Office 2010 SP 1 or newer is installed on the system, the VSTOR runtime
/// will be automagically configured as long as the .NET 4.0 runtime is present.
/// Office 2007 and Office 2010 without service pack need the VSTO runtime
/// redistributable. For details, see:
/// http://xltoolbox.sf.net/blog/2015/01/net-vsto-add-ins-getting-prerequisites-right
function NeedToInstallVstor(): boolean;
begin
	result := false; // Default for Office 2010 SP1 or newer
	if IsOffice2007Installed or IsOffice2010NoSpInstalled then
		result := not IsVstorInstalled;
end;

/// Checks if all required prerequisites are met, i.e. if the necessary
/// runtimes are installed on the system
function PrerequisitesAreMet(): boolean;
begin
	// Cache check result to avoid multiple registry lookups and log messages
	if not prerequisitesChecked then
	begin
		prerequisitesMet := IsNetInstalled and not NeedToInstallVstor;
		prerequisitesChecked := true;
	end;
	result := prerequisitesMet;
end;

/// Checks if a file exists and has a valid Sha1 sum.
function IsFileValid(file: string; expectedSha1: string): boolean;
var
	actualSha1: string;
begin
	try
		if FileExists(file) then
		begin
			actualSha1 := GetSHA1OfFile(file);
		end;
	finally
		result := actualSha1 = expectedSha1;
	end;
end;

/// Returns the path to the downloaded VSTO runtime installer.
function GetVstorInstallerPath(): string;
begin
	result := ExpandConstant('{%temp}\vstor_redist_40.exe');
end;

/// Returns the path to the downloaded .NET runtime installer.
function GetNetInstallerPath(): string;
begin
	result := ExpandConstant('{%temp}\dotNetFx40_Full_x86_x64.exe');
end;

/// Checks if the VSTO runtime redistributable setup file has already been
/// downloaded by comparing SHA1 checksums.
function IsVstorDownloaded(): boolean;
begin
	result := IsFileValid(GetVstorInstallerPath, '{#VSTORSHA1}');
end;

/// Checks if the .NET runtime setup file has already been
/// downloaded by comparing SHA1 checksums.
function IsNetDownloaded(): boolean;
begin
	result := IsFileValid(GetNetInstallerPath, '{#DOTNETSHA1}');
end;

/// Determines if the VSTO runtime needs to be downloaded.
/// This is not the case it the runtime is already installed,
/// or if there is a file with a valid Sha1 sum.
function NeedToDownloadVstor: boolean;
begin
	result := NeedToInstallVstor and not IsVstorDownloaded;
end;

/// Determines if the VSTO runtime needs to be downloaded.
/// This is not the case it the runtime is already installed,
/// or if there is a file with a valid Sha1 sum.
function NeedToDownloadNet: boolean;
begin
	result := not IsNetInstalled and not IsNetDownloaded;
end;

/// Determines whether or not a system-wide installation
/// is possible. This depends on whether the current user
/// is an administrator, and whether the hotfix KB976477
/// is present on the system if Excel 2007 is the only version
/// of Excel that is present (without that hotfix, Excel
/// 2007 does not load add-ins that are registered in the
/// HKLM hive).
function CanInstallSystemWide(): boolean;
begin
	if IsAdminLoggedOn then
	begin
		if IsOnlyExcel2007Installed then
		begin
			result := IsHotfixInstalled;
		end
		else
		begin
			result := true;
		end;
	end
	else
	begin
		result := false;
	end;
end;

procedure CreateDevelopmentInfoPage();
begin
	PageDevelopmentInfo := CreateInputOptionPage(wpWelcome,
		CustomMessage('DevVer'), CustomMessage('DevVerSubcaption'),
		CustomMessage('DevVerDesc'), False, False);
	PageDevelopmentInfo.Add(CustomMessage('DevVerCheckbox'));
	PageDevelopmentInfo.Values[0] := False;
end;

procedure CreateSingleOrAllUserPage();
begin
	PageSingleOrMultiUser := CreateInputOptionPage(wpLicense,
		CustomMessage('SingleOrMulti'), CustomMessage('SingleOrMultiSubcaption'),
		CustomMessage('SingleOrMultiDesc'), True, False);
	PageSingleOrMultiUser.Add(CustomMessage('SingleOrMultiSingle'));
	PageSingleOrMultiUser.Add(CustomMessage('SingleOrMultiAll'));
	if CanInstallSystemWide then
	begin
		PageSingleOrMultiUser.Values[1] := True;
	end
	else
	begin
		PageSingleOrMultiUser.Values[0] := True;
	end;
end;

procedure CreateCannotInstallPage();
begin
	PageCannotInstall := CreateInputOptionPage(wpWelcome,
		CustomMessage('CannotInstallCaption'),
		CustomMessage('CannotInstallDesc'),
		CustomMessage('CannotInstallMsg'), True, False);
	PageCannotInstall.Add(CustomMessage('CannotInstallCont'));
	PageCannotInstall.Add(CustomMessage('CannotInstallAbort'));
	PageCannotInstall.Values[1] := True;
end;

procedure CreateDownloadInfoPage();
var
	bytes: Int64;
	mib: Single;
	size: String;
begin
	if idpGetFilesSize(bytes) then
	begin
		mib := bytes / 1048576;
		size := Format('%.1f', [ mib ]);
	end
	else
	begin
		size := '[?]'
	end;
	PageDownloadInfo := CreateOutputMsgPage(PageSingleOrMultiUser.Id,
		CustomMessage('RequiredCaption'),
		CustomMessage('RequiredDesc'),
		Format(CustomMessage('RequiredMsg'), [idpFilesCount, size]));
end;

procedure CreateInstallInfoPage();
begin
	PageInstallInfo := CreateOutputMsgPage(PageDownloadInfo.Id,
		CustomMessage('InstallCaption'),
		CustomMessage('InstallDesc'),
		CustomMessage('InstallMsg'));
end;


function InitializeSetup(): boolean;
var
	minExcelInstalled: boolean;
	i: integer;
begin
	// Determine if Excel 2007 or newer is installed (absolute requirement
	// for this VSTO add-in). Excel 2007 ist version 12.0.
	for i := 12 to maxExcel do
	begin
		minExcelInstalled := minExcelInstalled or IsExcelVersionInstalled(i);
	end;

	if not minExcelInstalled then
	begin
		result := False;
		Log('Informing user that Excel 2007 or newer is required.');
		MsgBox(CustomMessage('Excel2007Required'), mbInformation, MB_OK);
	end
	else
	begin
		result := True;
	end
end;
	
procedure InitializeWizard();
begin
	CreateDevelopmentInfoPage;
	CreateSingleOrAllUserPage;
	if not PrerequisitesAreMet then
	begin
		Log('Not all prerequisites are met...');
		CreateCannotInstallPage;
		if NeedToDownloadNet then
		begin
			Log('Mark {#DOTNETURL} for download.');
			idpAddFileSize('{#DOTNETURL}', GetNetInstallerPath, 50449456);
		end;
		if NeedToDownloadVstor then
		begin
			Log('Mark {#VSTORURL} for download.');
			idpAddFileSize('{#VSTORURL}', GetVstorInstallerPath, 40123576);
		end;
		CreateDownloadInfoPage;
		CreateInstallInfoPage;
		idpDownloadAfter(PageDownloadInfo.Id);
	end;
end;

function ExecuteNetSetup(): boolean;
var
	exitCode: integer;
begin
	result := true;
	if not IsNetInstalled then
	begin
		if IsNetDownloaded then
		begin
			Log('Valid .NET runtime download found, installing.');
			Exec(GetNetInstallerPath, '/norestart',
				'', SW_SHOW, ewWaitUntilTerminated, exitCode);
			BringToFrontAndRestore;
			if not IsNetInstalled then
			begin
				MsgBox(CustomMessage('StillNotInstalled'), mbInformation, MB_OK);
				result := False;
			end;
		end
		else
		begin
			Log('No or invalid .NET runtime download found, will not install.');
			MsgBox(CustomMessage('DownloadNotValidated'), mbInformation, MB_OK);
			result := False;
		end;
	end; // not IsNetInstalled
end;

function ExecuteVstorSetup(): boolean;
var
	exitCode: integer;
begin
	result := true;
	if NeedToInstallVstor then
	begin
		if IsVstorDownloaded then
		begin
			Log('Valid VSTO runtime download found, installing.');
			Exec(GetVstorInstallerPath, '/norestart', '', SW_SHOW,
				ewWaitUntilTerminated, exitCode);
			BringToFrontAndRestore;
			if not IsVstorInstalled then
			begin
				MsgBox(CustomMessage('StillNotInstalled'), mbInformation, MB_OK);
				result := False;
			end;
		end
		else
		begin
			Log('No or invalid VSTO runtime download found, will not install.');
			MsgBox(CustomMessage('DownloadNotValidated'), mbInformation, MB_OK);
			result := False;
		end;
	end; // not IsVstorInstalled
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
	result := True;
	if not WizardSilent then
	begin
		if CurPageID = PageDevelopmentInfo.Id then
		begin
			if PageDevelopmentInfo.Values[0] = False then
			begin
				Log('Requesting user to acknowledge use of a developmental version.');
				MsgBox(CustomMessage('DevVerMsgBox'), mbInformation, MB_OK);
				result := False;
			end;
		end;
	end;
	
	if not PrerequisitesAreMet then
	begin
		// Abort the installation if any of the runtimes are missing, the user
		// is not an administrator, and requested to abort the installation.
		if CurPageID = PageCannotInstall.ID then
		begin
			if PageCannotInstall.Values[1] = true then
			begin
				WizardForm.Close;
				result := False;
			end
			else
			begin
				Log('Non-admin user continues although not all required runtimes are installed.');
			end;
		end;

		if CurPageID = PageInstallInfo.ID then
		begin
		  // Return true if installation succeeds (or no installation required)
			result := ExecuteNetSetup and ExecuteVstorSetup;
		end; 
	end; // not PrerequisitesAreMet
end;

/// Skips the folder selection, single/multi user, and ready pages for
/// normal users without power privileges.
/// This function also takes care of dynamically determining what wizard
/// pages to install, depending on the current system setup and whether
/// the current user is an administrator.
function ShouldSkipPage(PageID: Integer): Boolean;
begin
	result := False;
	
	if not PrerequisitesAreMet then
	begin
		// The PageDownloadCannotInstall will only have been initialized if
		// PrerequisitesAreMet returned false.
		if PageID = PageCannotInstall.ID then
		begin
			// Skip the warning if the user is an admin.
			result := IsAdminLoggedOn 
			if not result then
			begin
				Log('Warning user that required runtimes cannot be installed due to missing privileges');
			end;
		end;
		
		if PageID = PageDownloadInfo.ID then
		begin
			// Skip page informing about downloads if no files need to be downloaded.
			result := idpFilesCount = 0;
		end;
	
		if PageID = IDPForm.Page.ID then
		begin
			// Skip downloader plugin if there are no files to download.
			result := idpFilesCount = 0;
			if not result then
			begin
				Log('Beginning download of ' + IntToStr(idpFilesCount) + ' file(s).');
			end;
		end;
	end; // not PrerequisitesAreMet
	
	if PageID = PageSingleOrMultiUser.ID then
	begin
		if IsOnlyExcel2007Installed then
		begin
			Log('Only Excel 2007 appears to be installed on this system.');
			if IsHotfixInstalled then
			begin
				Log('Hotfix KB976477 found; can install for all users.');
			end
			else
			begin
				Log('Hotfix KB976477 not found; cannot install for all users.');
			end;
		end
		else
		begin
			Log('Excel 2010 or newer found on this system.');
		end;
		if CanInstallSystemWide then
		begin
			Log('Offer installation for all users.');
			result := False;
		end
		else
		begin
			Log('Offer single-user installation only.');
			result := True;
		end;
	end;
	
	if (PageID = wpSelectDir) or (PageID = wpReady) then
	begin
		// Do not show the pages to select the target directory, and the ready 
		// page if the user is not an admin.
		result := not ShouldShowDirPage;
	end
end;

/// Helper function that evaluates the custom PageSingleOrMultiUser page.
function IsMultiUserInstall(): Boolean;
begin
	result := PageSingleOrMultiUser.Values[1];
end;

/// Suggest an initial target directory depending on whether
/// the installer is run with admin privileges.
function SuggestInstallDir(Param: string): string;
var
	dir: string;
begin
	if CanInstallSystemWide then
	begin
		dir := ExpandConstant('{pf}');
	end
	else
	begin
    if IsZeroClient then
    begin
      dir := ExpandConstant('{localappdata}')
    end
    else
    begin
      dir := ExpandConstant('{userappdata}')
    end
	end;
	result := AddBackslash(dir) + 'Daniel''s XL Toolbox';
end;

/// Converts backslashes to forward slashes.
function ConvertSlash(Value: string): string;
begin
	StringChangeEx(Value, '\', '/', True);
	Result := Value;
end;
