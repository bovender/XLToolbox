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

[Languages]
Name: English; MessagesFile: compiler:Default.isl; 
Name: Deutsch; MessagesFile: compiler:Languages\German.isl; 
; Name: Brasileiro; MessagesFile: compiler:Languages\BrazilianPortuguese.isl;
; Name: Portugues; MessagesFile: compiler:Languages\Portuguese.isl;

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
English.DevVer=Developmental version
English.DevVerSubcaption=Please acknowledge that this is development in progress.
English.DevVerDesc=Please acknowledge that this is development in progress. Please note that this is a developmental version of the XL Toolbox. Expect things to work differently from what you expect to not work at all, to crash your system, and to change in any subsequent version.
English.DevVerCheckbox=I understand and expect things to not work and to crash at any time.
English.DevVerMsgBox=You must check the box to acknowledge that this is a development version in order to proceed.
English.SingleOrMulti=Single-user or system-wide install
English.SingleOrMultiSubcaption=Install for the current user only or for all users
English.SingleOrMultiDesc=Please indicate the scope of this installation:
English.SingleOrMultiSingle=Single user (only for me)
English.SingleOrMultiAll=All users (system-wide)
English.Excel2007Required=This add-in only works with Excel 2007 or later.

Deutsch.DevVer=Entwicklerversion
Deutsch.DevVerSubcaption=Bestätigen Sie, daß Sie die Entwicklerversion installieren wollen.
Deutsch.DevVerDesc=Beachten Sie bitte, daß es sich hierbei um eine Entwicklerversion handelt. Das Add-in kann sich anders verhalten, als Sie es erwarten, kann vielleicht Ihren Computer zum Absturz bringen, und kann in der nächsten Version ganz anders sein.
Deutsch.DevVerCheckbox=Ich verstehe, daß es sich um eine instabile Entwicklerversion handelt.
Deutsch.DevVerMsgBox=Sie müssen bestätigen, daß Sie die Entwicklerversion installieren wollen.
Deutsch.SingleOrMulti=Einzelner oder alle Benutzer
Deutsch.SingleOrMultiSubcaption=Geben Sie an, für wen die Installation sein soll
Deutsch.SingleOrMultiDesc=Bitte geben Sie an, ob die Toolbox nur für Sie oder für alle Benutzer installiert werden soll.
Deutsch.SingleOrMultiSingle=Ein Benutzer (nur für mich)
Deutsch.SingleOrMultiAll=Alle Benutzer (systemweit)
Deutsch.Excel2007Required=Diese Addin erfordert Excel 2007 oder eine neuere Version.

[Code]
var
	PageDevelopmentInfo: TInputOptionWizardPage;
	PageSingleOrMultiUser: TInputOptionWizardPage;

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
	if IsAdminLoggedOn then
	begin
		PageSingleOrMultiUser.Values[1] := True;
	end
	else
	begin
		PageSingleOrMultiUser.Values[0] := True;
	end;
end;

/// Checks if a given Excel version is installed
function IsExcelVersionInstalled(version: integer): boolean;
var key, wowNode: string;
var lookup1, lookup2: boolean;
begin
	key := 'SOFTWARE\Microsoft\Office\' + IntToStr(version) + '.0\Excel\InstallRoot';
	// The registry keys are located i different places depending
	// whether or not this is a 64-bit system.
	if IsWin64 then
	begin
		wowNode := 'Wow6432Node\';
	end;
	lookup1 := RegKeyExists(HKEY_LOCAL_MACHINE, wowNode + key);
	
	// If checking for version >= 14.0 ("2010"), which was the first version
	// that was produced in both 32-bit and 64-bit, on a 64-bit system we
	// also need to check a path without  'Wow6434Node'.
	if IsWin64 and (version >= 14) then
	begin
		lookup2 := RegKeyExists(HKEY_LOCAL_MACHINE, key);
	end;
	
	result := lookup1 or lookup2;
end;

function InitializeSetup(): boolean;
var
	minExcelInstalled: boolean;
	i: integer;
begin
	// The minimum required version of Excel is 2007 (12.0)
	for i := 12 to 24 do
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
