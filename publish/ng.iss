; Inno Setup script for Daniel's XL Toolbox
; (c) 2008-2014 Daniel Kraus
; GNU General Public License v2
; TODO: Include license file

[Setup]

#define SEMVER "7.0.0-alpha.1" ; Semantic version version
#define VER "7.0.0.0" ; The version in four-number format (Windows style)
#define YEAR "2014"
#define LOGFILE "INST-LOG.TXT"
#define REGKEY "Software\Microsoft\Office\Excel\Addins\XL Toolbox NG"
#define APPNAME "Daniel's XL Toolbox NG"
#define SLOGAN "Scientific add-in for Microsoft Excel."
#define RUNTIMEURL "http://vhost/vstor_redist.exe"
; "http://download.microsoft.com/download/2/E/9/2E9D2603-6D1F-4B12-BD37-DB1410B23597/vstor_redist.exe"
#define RUNTIMESHA1 "ad1dcc5325cb31754105c8c783995649e2208571"

; Specific AppID
AppId={{35AD3250-5F75-4C7D-BCE0-41377E280430}
                      
; Compiler info
OutputDir=release
OutputBaseFilename=XL_Toolbox_{#SEMVER}
Compression=lzma
InternalCompressLevel=max
SolidCompression=true
; LicenseFile=release\install license.rtf

; Application name, version, etc.
AppName={#APPNAME}
AppVerName={#APPNAME} {#SEMVER}
AppPublisher=Daniel Kraus
AppCopyright=2008-{#YEAR} Daniel Kraus
VersionInfoDescription={#SLOGAN}
VersionInfoVersion={#VER}
VersionInfoProductVersion={#VER}
VersionInfoTextVersion={#SEMVER}
VersionInfoProductTextVersion={#SEMVER}

AppPublisherURL=http://xltoolbox.sf.net
AppSupportURL=http://xltoolbox.sf.net/support.html
AppUpdatesURL=http://sf.net/projects/xltoolbox/files/latest

; Setup wizard
SetupIconFile=..\..\images\ico\xltoolbox.ico
WizardImageFile=..\..\images\installer\installbanner.bmp
WizardSmallImageFile=..\..\images\ico\icon-48.bmp
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
SetupLogging=true

; Uninstall
; UninstallDisplayIcon={app}\img\xltoolbox.ico
UninstallFilesDir={app}\uninstall
	
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
Source: ..\XLToolbox\bin\Release\*; DestDir: {app}; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Registry]
; Keys for single-user install (HKCU)
Check: not IsMultiUserInstall; ValueName: Description; ValueData: {#SLOGAN}; ValueType: string; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: FriendlyName; ValueData: {#APPNAME}; ValueType: string; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: LoadBehavior; ValueData: 3; ValueType: dword; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: Warmup; ValueData: 1; ValueType: dword; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: Manifest; ValueData: file:///{code:ConvertSlash|{app}}/XLToolbox.vsto|vstolocal; ValueType: string; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey

; Same keys again, this time for multi-user install (HKLM)
Check: IsMultiUserInstall; ValueName: Description; ValueData: {#SLOGAN}; ValueType: string; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: FriendlyName; ValueData: {#APPNAME}; ValueType: string; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: LoadBehavior; ValueData: 3; ValueType: dword; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: Warmup; ValueData: 1; ValueType: dword; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: Manifest; ValueData: file:///{code:ConvertSlash|{app}}/XLToolbox.vsto|vstolocal; ValueType: string; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey

[Run]
; Filename: http://xltoolbox.sourceforge.net/welcome.html; Flags: shellexec nowait; Check: ShowWelcomePageInBrowser

[UninstallRun]
; Filename: http://xltoolbox.sourceforge.net/uninstall.html; Flags: shellexec nowait

[CustomMessages]
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
en.ClrDownloadCaption=Runtime files required
en.ClrDownloadDesc=The required Visual Studio Tools for Office (VSTO) 4.0 runtime files were not found on your system.
en.ClrDownloadMsg=Click 'Next' to start downloading the installer file (about 38 MB) from the Microsoft servers.
en.ClrInstallCaption=Runtime files downloaded
en.ClrInstallDesc=The Visual Studio Tools for Office (VSTO) 4.0 runtime files are ready to install.
en.ClrInstallMsg=Click 'Next' to beginn the installation of the runtime files.
en.ClrCannotInstallCaption=Administrator privileges required
en.ClrCannotInstallDesc=You do not have the administrative rights to install the required Visual Studio for Office (VSTO) 4.0 runtime files.
en.ClrCannotInstallMsg=You may continue the installation, but the XL Toolbox won't start unless the required VSTO runtime files are installed by an administrator.
en.ClrCannotInstallCont=Continue anyway, although it won't work without the runtime files
en.ClrCannotInstallAbort=Abort the installation (come back when the admin has installed the files)
en.ClrStillNotInstalled=The required VSTO runtime is still not installed. Setup cannot continue. You may try again, or abort the XL Toolbox installation.
en.ClrNotValidated=The downloaded file has unexpected content. It may have not been downloaded correctly, or someone might have hampered with it. You may click 'Back' and then 'Next' to download it again.

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
de.ClrDownloadCaption=Laufzeitdateien erforderlich
de.ClrDownloadDesc=Die benötigten Laufzeitdateien der Visual Studio Tools for Office (VSTO) 4.0 wurden nicht auf Ihrem System gefunden.
de.ClrDownloadMsg=Klicken Sie 'Weiter', um mit dem Download der Installationsdatei von den Microsoft-Servern zu beginnen.
de.ClrInstallCaption=Laufzeitdateien wurden heruntergeladen
de.ClrInstallDesc=Die Laufzeitdateien der Visual Studio Tools for Office (VSTO) 4.0 können jetzt installiert werden.
de.ClrInstallMsg=Klicken Sie 'Weiter', um die Installation der Laufzeitdateien zu starten.
de.ClrCannotInstallCaption=Administratorrechte benötigt
de.ClrCannotInstallDesc=Sie sind kein Admin und daher nicht autorisiert, die erforderlichen Laufzeitdateien von Visual Studio for Office (VSTO) 4.0 zu installieren.
de.ClrCannotInstallMsg=Sie können mit der Installation fortfahren, aber die Toolbox wird nicht starten, solange die VSTO-Laufzeitdateien nicht von einem Admin installiert wurden.
de.ClrCannotInstallCont=Trotzdem installieren, obwohl es nicht funktionieren wird
de.ClrCannotInstallAbort=Installation abbrechen
de.ClrStillNotInstalled=Die benötigten VSTO-Laufzeitdateien sind immer noch nicht installiert. Sie können es noch einmal versuchen oder die Installation der XL Toolbox abbrechen.
de.ClrNotValidated=Die heruntergeladene Datei enthält unerwartete Daten. Vielleicht wurde sie nicht korrekt heruntergeladen, oder sie wurde von jemandem manipuliert. Sie können 'Zurück' und dann 'Weiter' klicken, um sie erneut herunterzuladen.

[Code]
const
	maxExcel = 24; //< highest Excel version number to check for.
	
var
	PageDevelopmentInfo: TInputOptionWizardPage;
	PageSingleOrMultiUser: TInputOptionWizardPage;
	PageClrCannotInstall: TInputOptionWizardPage;
	PageClrDownloadInfo: TOutputMsgWizardPage;
	PageClrInstallInfo: TOutputMsgWizardPage;

/// Returns the path for the Wow6432Node registry tree if the current operating
/// system is 64-bit.
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

/// Checks if the CLR is installed. This is relevant if only
/// Excel 2007 is installed. Since Office 2010, the CLR is
/// automatically included.
/// The presence of the 4.0 CLR is indicated by the presence one of
/// four possible registry keys (cf. http://stackoverflow.com/a/15311013/270712):
/// HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4 (32-bit, VSTO installed from Office 2010 installation)
/// HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R (32-bit, VSTO installed from redistributable)
/// HKLM\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4 (64-bit, VSTO installed from Office 2010 installation)
/// HKLM\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R (64-bit, VSTO installed from redistributable)
function IsClrInstalled(): boolean;
var
	software, clrPath, wowNode: string;
begin
	software := 'SOFTWARE\';
	clrPath := 'Microsoft\VSTO Runtime Setup\v4';
	wowNode := GetWowNode;
	result := RegKeyExists(HKEY_LOCAL_MACHINE, software + wowNode + clrPath) or
		RegKeyExists(HKEY_LOCAL_MACHINE, software + wowNode + clrPath + 'R');
end;

/// Returns the path to the downloaded VSTO runtime installer.
function GetClrInstallerPath(): string;
begin
	result := ExpandConstant('{%temp}\vstor_redist_40.exe');
end;

/// Checks if the CLR redistributable setup file has already been
/// downloaded by comparing SHA1 checksums.
function IsClrDownloaded(): boolean;
var
	downloadedSha1: string;
begin
	try
		if FileExists(GetClrInstallerPath) then
		begin
			downloadedSha1 := GetSHA1OfFile(GetClrInstallerPath);
		end;
	finally
		result := downloadedSha1 = '{#RUNTIMESHA1}';
	end;
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
	PageSingleOrMultiUser := CreateInputOptionPage(PageDevelopmentInfo.ID,
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

procedure CreateClrCannotInstallPage();
begin
	PageClrCannotInstall := CreateInputOptionPage(wpWelcome,
		CustomMessage('ClrCannotInstallCaption'),
		CustomMessage('ClrCannotInstallDesc'),
		CustomMessage('ClrCannotInstallMsg'), True, False);
	PageClrCannotInstall.Add(CustomMessage('ClrCannotInstallCont'));
	PageClrCannotInstall.Add(CustomMessage('ClrCannotInstallAbort'));
	PageClrCannotInstall.Values[1] := True;
end;

procedure CreateClrDownloadInfoPage();
begin
	PageClrDownloadInfo := CreateOutputMsgPage(PageSingleOrMultiUser.Id,
		CustomMessage('ClrDownloadCaption'),
		CustomMessage('ClrDownloadDesc'),
		CustomMessage('ClrDownloadMsg'));
end;

procedure CreateClrInstallInfoPage();
begin
	PageClrInstallInfo := CreateOutputMsgPage(PageClrDownloadInfo.Id,
		CustomMessage('ClrInstallCaption'),
		CustomMessage('ClrInstallDesc'),
		CustomMessage('ClrInstallMsg'));
end;

function InitializeSetup(): boolean;
var
	minExcelInstalled: boolean;
	i: integer;
begin
	// The minimum required version of Excel is 2007 (12.0)
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
	CreateClrCannotInstallPage;
	CreateClrDownloadInfoPage;
	CreateClrInstallInfoPage;
	idpAddFile('{#RUNTIMEURL}', GetClrInstallerPath);
	idpDownloadAfter(PageClrDownloadInfo.Id);
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
	exitCode: integer;
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
	
	// Abort the installation if the VSTO runtime is missing, the user
	// is not an administrator, and requested to abort the installation.
	if CurPageID = PageClrCannotInstall.ID then
	begin
		if PageClrCannotInstall.Values[1] = true then
		begin
			WizardForm.Close;
			result := False;
		end
		else
			Log('Non-admin user continues although VSTO runtime is not installed.');
		begin
		end;
	end;

	// Trigger the VSTO runtime installation
	if CurPageID = PageClrInstallInfo.ID then
	begin
		if IsClrDownloaded then
		begin
			Log('Valid VSTO runtime download found, installing.');
			Exec(GetClrInstallerPath, '', '', SW_SHOW, ewWaitUntilTerminated, exitCode);
			BringToFrontAndRestore;
			if not IsClrInstalled then
			begin
				MsgBox(CustomMessage('ClrStillNotInstalled'), mbInformation, MB_OK);
				result := False;
			end;
		end
		else
		begin
			Log('Invalid VSTO runtime download found, will not install.');
			MsgBox(CustomMessage('ClrNotValidated'), mbInformation, MB_OK);
			result := False;
		end;
	end;
end;

/// Skips the folder selection, single/multi user, and ready pages for
/// normal users without power privileges.
/// This function also takes care of dynamically determining what wizard
/// pages to install, depending on the current system setup and whether
/// the current user is an administrator.
function ShouldSkipPage(PageID: Integer): Boolean;
begin
	result := False;

	if (PageID = PageClrInstallInfo.ID) or (PageID = PageClrDownloadInfo.ID) then
	begin
		// Skip the pages to download and install the VSTO runtime.
		result := IsClrInstalled;
	end;
	
	if PageID = PageClrCannotInstall.ID then
	begin
		// Only warn the user about the missing VSTO runtime that cannot be 
		// installed if the runtime is really missing and the user is not
		// an admin.
		if not IsClrInstalled then
		begin
			// Skip the page if the user is an admin.
			if IsAdminLoggedOn then
			begin
				result := True;
			end
			else
			begin
				Log('Warning user that VSTO runtime cannot be installed due to missing privileges');
				result := False;
			end;
		end
		else
		begin
			// Do not show the warning if the runtime is installed, regardless
			// if the user is an admin or not.
			result := True;
		end;
	end;
	
	if PageID = IDPForm.Page.ID then
	begin
		if IsClrInstalled then
		begin
			// Skip the download page if the VSTO runtime is already installed.
			result := True;
			Log('VSTO runtime is already installed on this system.');
		end
		else
		begin
			Log('VSTO runtime is not installed on this system.');

			// Skip the download page if the runtime installer has already been
			// downloaded.
			if IsClrDownloaded then
			begin
				Log('VSTO runtime installer found, skipping download.');
				result := True;
			end
			else
			begin
				Log('Downloading VSTO runtime installer.');
				Log('URL: {#RUNTIMEURL}');
				result := False;
			end;
		end;
	end;
	
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
		result := not IsAdminLoggedOn;
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
	if IsAdminLoggedOn then
	begin
		dir := ExpandConstant('{pf}');
	end
	else
	begin
		dir := ExpandConstant('{userappdata}');
	end;
	result := AddBackslash(dir) + 'Daniel''s XL Toolbox';
end;

/// Converts backslashes to forward slashes.
function ConvertSlash(Value: string): string;
begin
	StringChangeEx(Value, '\', '/', True);
	Result := Value;
end;
