; Inno Setup script for Daniel's XL Toolbox
; (c) 2008-2014 Daniel Kraus
; GNU General Public License v2
; TODO: Include license file

[Setup]

#define SEMVER "7.0.0-alpha.1" ; Semantic version version
#define VER "7.0.0.0" ; The version in four-number format (Windows style)
#define YEAR "2014"
#define LOGFILE "INST-LOG.TXT"
#define REGKEY "Software\Microsoft\Office\Excel\Addins\XL Toolbox"
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
de.Excel2007Required=Daniel's XL Toolbox NG erfordert 

[Code]
const
	maxExcel = 24; //< highest Excel version number to check for.
	
var
	PageDevelopmentInfo: TInputOptionWizardPage;
	PageSingleOrMultiUser: TInputOptionWizardPage;

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
		lookup2 := RegKeyExists(HKEY_LOCAL_MACHINE, key);
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
function IsCLRInstalled(): boolean;
var
	software, clrPath, wowNode: string;
begin
	software := 'SOFTWARE\';
	clrPath := 'Microsoft\VSTO Runtime Setup\v4';
	wowNode := GetWowNode;
	result := RegKeyExists(HKEY_LOCAL_MACHINE, software + wowNode + clrPath) or
		RegKeyExists(HKEY_LOCAL_MACHINE, software + wowNode + clrPath + 'R');
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
		MsgBox(CustomMessage('Excel2007Required'), mbInformation, MB_OK);
	end
	else
	begin
		result := True;
	end
end;
	
procedure InitializeWizard();
begin
	CreateDevelopmentInfoPage();
	CreateSingleOrAllUserPage();
	
	idpAddFile('{#RUNTIMEURL}', ExpandConstant('{tmp}\xltb_vstor.exe'));
	idpDownloadAfter(wpWelcome);
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
	result := True;
	if not WizardSilent then
	begin
		if CurPageID = PageDevelopmentInfo.ID then
		begin
			if PageDevelopmentInfo.Values[0] = False then
			begin
				MsgBox(CustomMessage('DevVerMsgBox'), mbInformation, MB_OK);
				result := False;
			end;
		end;
	end;
end;

/// Skips the folder selection, single/multi user, and ready pages for
/// normal users without power privileges..
function ShouldSkipPage(PageID: Integer): Boolean;
begin
	result := False;
	if not IsAdminLoggedOn then
	begin
		if (PageID = wpSelectDir) or (PageID = wpReady) or
			(PageID = PageSingleOrMultiUser.ID) then
		begin
			result := True;
		end
	end;
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
