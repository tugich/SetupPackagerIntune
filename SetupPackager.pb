
;- Notes
;-- Default Window Height is 520px

;- Compiler
EnableExplicit
UsePNGImageDecoder()

CompilerIf #PB_Compiler_OS = #PB_OS_Windows
  Macro GetParentDirectory(Path)
    GetPathPart(RTrim(Path, "\"))
  EndMacro
CompilerElse
  Macro GetParentDirectory(Path)
    GetPathPart(RTrim(Path, "/"))
  EndMacro
CompilerEndIf

;- Variables

;-- Window
Global Event = #Null, Quit = #False
Global Image_DropFolder, Combo_InstallerFile, Image_GreenCheck, 
       Hyperlink_PackageFolder, Hyperlink_PSADT_Folder, Checkbox_CreateSubFolders,
       Button_CreatePackage, Button_RunInstaller, Button_PSAppDeployToolkit, Button_MSI, Button_DebugPSADT,
       Text_Introduction

;-- Packaging
Global InstallerFile.s, InstallerPath.s
Global DropFolderPath.s, PackagePath.s, PSADT_Path.s
Global CreationWindowActive = #False, PSADT_CloseApps.s

;-- Structures
Structure MSI_Information
  Publisher.s
  Name.s
  Version.s
  Productcode.s
EndStructure

Global Details.MSI_Information
Details\Publisher = ""
Details\Name = ""
Details\Version = ""
Details\Productcode = ""

Structure InstalledApp
  Publisher.s
  DisplayName.s
  DisplayVersion.s
EndStructure
Global NewList InstalledApps.InstalledApp()

;-- Enumerations
Enumeration KeyboardShortcuts
  #SearchWindow_EnterPressed
EndEnumeration

;-- Executables
Global IntuneWinAppUtil.s = "IntuneWinAppUtil.exe"
Global PEiD.s = "ThirdParty/PEiD/PEiD.exe"

;- Libraries
XIncludeFile "Libraries/Registry.pbi"

;- Forms
XIncludeFile "Forms/MainWindow.pbf"
XIncludeFile "Forms/AboutWindow.pbf"
XIncludeFile "Forms/DetailsWindow_MSI.pbf"
XIncludeFile "Forms/DebuggerWindow_PSADT.pbf"
XIncludeFile "Forms/CreationWindow_PSADT.pbf"
XIncludeFile "Forms/SearchWindow_InstalledApps.pbf"

;- Functions
Procedure.s GetExePath()
  Protected Path.s = GetPathPart(ProgramFilename())
  If Path = GetTemporaryDirectory()
    Path = GetCurrentDirectory()
  EndIf
  ProcedureReturn Path
EndProcedure

Procedure Registry_UpdateInstalledApps(Path.s = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
  Define InstalledApps_Count, i
  InstalledApps_Count = Registry::CountSubKeys(#HKEY_LOCAL_MACHINE, Path)
  
  For i = 0 To InstalledApps_Count - 1
    Define App_ID.s, App_Publisher.s, App_DisplayName.s, App_DisplayVersion.s, App_Result.s
    App_ID = Registry::ListSubKey(#HKEY_LOCAL_MACHINE, Path, i)
    App_Publisher = Registry::ReadValue(#HKEY_LOCAL_MACHINE, Path + "\" + App_ID, "Publisher")
    App_DisplayName = Registry::ReadValue(#HKEY_LOCAL_MACHINE, Path + "\" + App_ID, "DisplayName")
    App_DisplayVersion = Registry::ReadValue(#HKEY_LOCAL_MACHINE, Path + "\" + App_ID, "DisplayVersion")
    
    If App_DisplayName <> ""
      AddElement(InstalledApps())
      InstalledApps()\Publisher = App_Publisher
      InstalledApps()\DisplayName = App_DisplayName
      InstalledApps()\DisplayVersion = App_DisplayVersion
    EndIf
  Next
EndProcedure

Procedure DownloadIntuneWinAppUtil(Parameter)
  
  Protected DestinationPath.s = GetTemporaryDirectory() + "IntuneWinAppUtil.exe"

  If ReceiveHTTPFile("https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe", DestinationPath)
    ;Debug "Downloaded successfully: " + DestinationPath
    IntuneWinAppUtil = DestinationPath
  Else
    Debug "Failed to download IntuneWinAppUtil!"
    MessageRequester("Error", "Unfortunately, something went wrong during the download of IntuneWinAppUtil. Please check your internet connection.", #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
  EndIf
  
EndProcedure

Procedure PSADT_ReadInformations(EventType)
  ;Debug DropFolderPath
  
  Protected FileIn, FileOut, sLine.s
  FileIn = ReadFile(#PB_Any, DropFolderPath + "\Deploy-Application.ps1", #PB_File_SharedRead)
  
  If FileIn
    ; Read each line from the input file, replace text, and write to the output file
    While Not Eof(FileIn)
      sLine = ReadString(FileIn)
      sLine = LTrim(sLine)

      If FindString(sLine, "[String]$appVendor")
        sLine = RemoveString(sLine, "[String]$appVendor =")
        sLine = RemoveString(sLine, "[String]$appVendor=")
        sLine = LTrim(sLine)
        sLine = LTrim(sLine, "'")
        sLine = RTrim(sLine, "'")
        SetGadgetText(PD_String_Publisher, sLine)
      EndIf
      
      If FindString(sLine, "[String]$appName")
        sLine = RemoveString(sLine, "[String]$appName =")
        sLine = RemoveString(sLine, "[String]$appName=")
        sLine = LTrim(sLine)
        sLine = LTrim(sLine, "'")
        sLine = RTrim(sLine, "'")
        SetGadgetText(PD_String_ProductName, sLine)
      EndIf
      
      If FindString(sLine, "[String]$appVersion")
        sLine = RemoveString(sLine, "[String]$appVersion =")
        sLine = RemoveString(sLine, "[String]$appVersion=")
        sLine = LTrim(sLine)
        sLine = LTrim(sLine, "'")
        sLine = RTrim(sLine, "'")
        SetGadgetText(PD_String_Version, sLine)
      EndIf
      
    Wend
    
    CloseFile(FileIn)
  EndIf
EndProcedure

Procedure ShowWebsite(EventType)
  RunProgram("https://blog.tugi.ch/setup-packager-intune", "", "")
EndProcedure

Procedure ShowAboutWindow(EventType)
  If IsWindow(AboutWindow)
    HideWindow(AboutWindow, #False)
    SetActiveWindow(AboutWindow)
  Else
    OpenAboutWindow()
  EndIf
EndProcedure

Procedure ShowDebugWindow_PSADT(EventType)
  If IsWindow(DebuggerWindow_PSADT)
    HideWindow(DebuggerWindow_PSADT, #False)
    SetActiveWindow(DebuggerWindow_PSADT)
  Else
    OpenDebuggerWindow_PSADT()
  EndIf

  CreateThread(@PSADT_ReadInformations(), 0)
EndProcedure

Procedure ShowCreationWindow(EventType)
  If IsWindow(CreationWindow_PSADT)
    HideWindow(CreationWindow_PSADT, #False)
    SetActiveWindow(CreationWindow_PSADT)
  Else
    OpenCreationWindow_PSADT()
  EndIf
EndProcedure

Procedure ShowSearchWindow(EventType)
  If IsWindow(SearchWindow_InstalledApps)
    HideWindow(SearchWindow_InstalledApps, #False)
    SetActiveWindow(SearchWindow_InstalledApps)
  Else
    OpenSearchWindow_InstalledApps()
  EndIf
  
  ; Shortcut for Enter
  AddKeyboardShortcut(SearchWindow_InstalledApps, #PB_Shortcut_Return, #SearchWindow_EnterPressed)
  
  ; Remove all gadget items in the listview
  ClearGadgetItems(SW_ListView_Result)
  
  ; Clear my list
  ClearList(InstalledApps())
  
  ; Read installed apps from Windows Registry
  Registry_UpdateInstalledApps("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
  Registry_UpdateInstalledApps("SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
  
  ; Update Listview
  ForEach InstalledApps()
    AddGadgetItem(SW_ListView_Result, -1, InstalledApps()\Publisher + " - " + InstalledApps()\DisplayName + " (" + InstalledApps()\DisplayVersion + ")")
  Next
EndProcedure

Procedure CloseAboutWindow(EventType)
  HideWindow(AboutWindow, #True)
  SetActiveWindow(MainWindow)
EndProcedure

Procedure CloseDebuggerWindow(EventType)
  HideWindow(DebuggerWindow_PSADT, #True)
  SetActiveWindow(MainWindow)
EndProcedure

Procedure CloseDetailsWindow(EventType)
  HideWindow(DetailsWindow_MSI, #True)
  SetActiveWindow(MainWindow)
EndProcedure

Procedure CloseCreationWindow(EventType)
  If IsWindow(CreationWindow_PSADT)
    HideWindow(CreationWindow_PSADT, #True)
  EndIf
  
  SetActiveWindow(MainWindow)
  CreationWindowActive = #False
  
  DisableGadget(Button_PSAppDeployToolkit, #False)
  DisableGadget(Button_RunInstaller, #False)
  DisableGadget(Button_CreatePackage, #False)
  DisableGadget(Combo_InstallerFile, #False)
EndProcedure

Procedure CloseSearchWindow(EventType)
  HideWindow(SearchWindow_InstalledApps, #True)
  
  If CreationWindowActive
    SetActiveWindow(CreationWindow_PSADT)
  Else
    SetActiveWindow(MainWindow)
  EndIf
EndProcedure

Procedure SearchInstalledApp(EventType)
  Protected SearchText.s = GetGadgetText(SW_String_Search)
  Protected NewList SearchResult.InstalledApp()
  
  If SearchText <> ""
    ; Remove all gadget items in the listview
    ClearGadgetItems(SW_ListView_Result)
    
    ; Build search text
    SearchText = RemoveString(SearchText, " ")
    SearchText = LCase(SearchText)
    Debug "Searching with: " + SearchText
    
    ForEach InstalledApps()
      Define SearchComparison.s
      
      SearchComparison = InstalledApps()\Publisher + InstalledApps()\DisplayName
      SearchComparison = RemoveString(SearchComparison, " ")
      SearchComparison = LCase(SearchComparison)
      
      If FindString(SearchComparison, SearchText, 0, #PB_String_NoCase)
        AddGadgetItem(SW_ListView_Result, -1, InstalledApps()\Publisher + " - " + InstalledApps()\DisplayName + " (" + InstalledApps()\DisplayVersion + ")")
      EndIf
    Next
  Else
    ShowSearchWindow(0)
  EndIf
EndProcedure

Procedure.s PowerShell_MsiProductCode(FilePath.s)
  Protected Compiler = #Null
  Protected Output$ = ""
  Protected Exitcode$ = ""
  Protected PSExitcode.i = 1
  Protected Productcode.s = ""

  Compiler = RunProgram("powershell.exe", 
                        "-NoProfile -NoLogo -WindowStyle Hidden -File .\Scripts\Read-Msi.ps1 -FilePath " + Chr(34) + FilePath + Chr(34) + "  -ExecutionPolicy Bypass", 
                        "", 
                        #PB_Program_Open | #PB_Program_Hide | #PB_Program_Read)
  Output$ = ""
  
  If Compiler
    While ProgramRunning(Compiler)
      If AvailableProgramOutput(Compiler)
        Output$ = ReadProgramString(Compiler)
        
        If FindString(Output$, "Productcode:")
          Productcode = RemoveString(Output$, "Productcode: ")
        EndIf
      EndIf
    Wend

    PSExitcode = ProgramExitCode(Compiler)
    CloseProgram(Compiler)
  EndIf
  
  If (PSExitcode = 0)
    ProcedureReturn Productcode
  Else
    MessageRequester("PowerShell Error", Output$, #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
  EndIf
EndProcedure

Procedure Batch_IntuneUpload(EventType)
  
  Protected Compiler = #Null
  Protected Output$ = ""
  Protected Exitcode$ = ""
  Protected PSExitcode.i = 1
  
  Protected IntuneWinFile.s = PackagePath + "\" + GetFilePart(InstallerFile, #PB_FileSystem_NoExtension) + ".intunewin" 
  Protected TenantDomain.s = InputRequester("Upload to Microsoft Intune", "Please enter your Microsoft Tenant Domain:", "")
  
  ; Check Microsoft Tenant Domain
  If Trim(TenantDomain) = ""
    MessageRequester("Upload to Intune failed", "Please provide valid tenant domain.", #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
    ProcedureReturn #False
  EndIf
  
  Protected AppName.s = InputRequester("Upload to Microsoft Intune", "Please enter the display name for the app:", "My App 1.0 (x64)")
  
  ; Check App Display Name
  If Trim(AppName) = ""
    MessageRequester("Upload to Intune failed", "Please provide valid display name for the app.", #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
    ProcedureReturn #False
  EndIf
  
  DisableGadget(Button_UploadIntune, #True)
  SetGadgetText(Button_UploadIntune, "Please wait...")

  Compiler = RunProgram("Scripts\UploadTo-Intune.bat", 
                        TenantDomain + " " + Chr(34) + IntuneWinFile + Chr(34) + " " + Chr(34) + AppName + Chr(34), 
                        "", 
                        #PB_Program_Open | #PB_Program_Read)
  Output$ = ""
  
  If Compiler
    While ProgramRunning(Compiler)
      If AvailableProgramOutput(Compiler)
        Output$ + ReadProgramString(Compiler) + Chr(13)
      EndIf
    Wend

    PSExitcode = ProgramExitCode(Compiler)
    CloseProgram(Compiler)
  EndIf
  
  ; Enable Button
  DisableGadget(Button_UploadIntune, #False)
  
  ; Check Exit code
  If (PSExitcode = 0)
    SetGadgetText(Button_UploadIntune, "Uploaded successfully")
    ProcedureReturn #True
  Else
    MessageRequester("PowerShell Error", Output$, #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
    SetGadgetText(Button_UploadIntune, "Upload failed - Retry")
    ProcedureReturn #False
  EndIf
  
EndProcedure

Procedure Powershell_MsiDetails(Parameter)
  Protected InstallerFile.s = GetGadgetText(Combo_InstallerFile)
  Protected FilePath.s = DropFolderPath + "\" + InstallerFile
  Protected Compiler = #Null
  Protected Output$ = ""
  Protected Exitcode$ = ""
  Protected PSExitcode.i = 1
  Protected Details.MSI_Information
  
  Compiler = RunProgram("powershell.exe", 
                        "-NoProfile -NoLogo -WindowStyle Hidden -File .\Scripts\Read-Msi.ps1 -FilePath " + Chr(34) + FilePath + Chr(34) + "  -ExecutionPolicy Bypass", 
                        "", 
                        #PB_Program_Open | #PB_Program_Hide | #PB_Program_Read)
  Output$ = ""
  
  If Compiler
    While ProgramRunning(Compiler)
      If AvailableProgramOutput(Compiler)
        Output$ = ReadProgramString(Compiler)
        
        If FindString(Output$, "Productmanufacturer:")
          Details\Publisher = RemoveString(Output$, "Productmanufacturer: ")
        ElseIf FindString(Output$, "Productname:")
          Details\Name = RemoveString(Output$, "Productname: ")
        ElseIf FindString(Output$, "Productcode:")
          Details\Productcode = RemoveString(Output$, "Productcode: ")
        ElseIf FindString(Output$, "Productversion:")
          Details\Version = RemoveString(Output$, "Productversion: ")
        EndIf
      EndIf
    Wend

    PSExitcode = ProgramExitCode(Compiler)
    CloseProgram(Compiler)
  EndIf
  
  If (PSExitcode = 0)
    SetGadgetText(DW_String_Name, Details\Name)
    SetGadgetText(DW_String_Version, Details\Version)
    SetGadgetText(DW_String_Publisher, Details\Publisher)
    AddGadgetItem(DW_ListIcon_ProductIDS, -1, Details\Productcode)
    DisableGadget(DW_Button_CopyProductcode, #False)
  Else
    MessageRequester("PowerShell Error", Output$, #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
  EndIf
  
  DisableGadget(Button_MSI, #False)
EndProcedure

Procedure Powershell_MsiDetails2()
  Protected InstallerFile.s = GetGadgetText(Combo_InstallerFile)
  Protected FilePath.s = DropFolderPath + "\" + InstallerFile
  Protected Compiler = #Null
  Protected Output$ = ""
  Protected Exitcode$ = ""
  Protected PSExitcode.i = 1
  
  Compiler = RunProgram("powershell.exe", 
                        "-NoProfile -NoLogo -WindowStyle Hidden -File .\Scripts\Read-Msi.ps1 -FilePath " + Chr(34) + FilePath + Chr(34) + "  -ExecutionPolicy Bypass", 
                        "", 
                        #PB_Program_Open | #PB_Program_Hide | #PB_Program_Read)
  Output$ = ""
  
  If Compiler
    While ProgramRunning(Compiler)
      If AvailableProgramOutput(Compiler)
        Output$ = ReadProgramString(Compiler)
        
        If FindString(Output$, "Productmanufacturer:")
          Details\Publisher = RemoveString(Output$, "Productmanufacturer: ")
        ElseIf FindString(Output$, "Productname:")
          Details\Name = RemoveString(Output$, "Productname: ")
        ElseIf FindString(Output$, "Productcode:")
          Details\Productcode = RemoveString(Output$, "Productcode: ")
        ElseIf FindString(Output$, "Productversion:")
          Details\Version = RemoveString(Output$, "Productversion: ")
        EndIf
      EndIf
    Wend

    PSExitcode = ProgramExitCode(Compiler)
    CloseProgram(Compiler)
  EndIf
  
  If (PSExitcode = 0)
    ProcedureReturn Details
  Else
    MessageRequester("PowerShell Error", Output$, #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
  EndIf
EndProcedure

Procedure ShowMsiProductCode(EventType)
  
  ;Debug InstallerPath.s
  
  If IsWindow(DetailsWindow_MSI)
    HideWindow(DetailsWindow_MSI, #False)
    SetActiveWindow(DetailsWindow_MSI)
  Else
    OpenDetailsWindow_MSI()
  EndIf
  
  DisableGadget(Button_MSI, #True)
  ClearGadgetItems(DW_ListIcon_ProductIDS)
  HideGadget(Image_CopyProductcode_Check, #True)
  CreateThread(@Powershell_MsiDetails(), 0)
  
  ;SetGadgetText(DW_String_Name, PowerShell_MsiProductCode(DropFolderPath + "\" + InstallerFile))
  
EndProcedure

Procedure DropFolder(Folder.s)
  
  ClearGadgetItems(Combo_InstallerFile)
  SetGadgetText(Text_Introduction, Folder)
  DisableGadget(Button_CreatePackage, #False)
  DisableGadget(Combo_InstallerFile, #False)
  DisableGadget(Button_UploadIntune, #True)
  DisableGadget(Button_PSAppDeployToolkit, #True)
  DisableGadget(Button_RunInstaller, #True)
  DisableGadget(Button_PEiD, #True)
  DisableGadget(Button_MSI, #True)
  HideGadget(Button_SelectFolder, #True)
  HideGadget(Button_DebugPSADT, #True)
  HideGadget(Hyperlink_PackageFolder, #True)
  HideGadget(Hyperlink_PSADT_Folder, #True)
  HideGadget(Image_GreenCheck, #True)
  SetGadgetText(Button_UploadIntune, "Upload to Intune")
  ResizeWindow(MainWindow, WindowX(MainWindow), WindowY(MainWindow), WindowWidth(MainWindow), 520)
  
  ; Close Windows
  If IsWindow(DetailsWindow_MSI)
    CloseAboutWindow(0)
  EndIf
  
  If IsWindow(DebuggerWindow_PSADT)
    CloseDebuggerWindow(0)
  EndIf
  
  ; List files
  If ExamineDirectory(0, Folder, "*.*")
    While NextDirectoryEntry(0)
      Protected Type$, Size$
      
      If DirectoryEntryType(0) = #PB_DirectoryEntry_File And GetExtensionPart(DirectoryEntryName(0)) <> "intunewin"
        AddGadgetItem(Combo_InstallerFile, -1, DirectoryEntryName(0))
      EndIf
    Wend
    FinishDirectory(0)
  EndIf
  
  ; Check if PSADT project
  If ExamineDirectory(0, Folder, "*.*")
    While NextDirectoryEntry(0)
      If DirectoryEntryType(0) = #PB_DirectoryEntry_File And DirectoryEntryName(0) = "Deploy-Application.exe"
        HideGadget(Button_DebugPSADT, #False)
        SetGadgetText(Combo_InstallerFile, "Deploy-Application.exe")
        MyInstallerFile(#PB_EventType_Change)
        Break
      EndIf
    Wend
    FinishDirectory(0)
  EndIf
  
EndProcedure

Procedure SelectFolder(EventType)
  
  If EventType <> #PB_EventType_LeftClick
    ProcedureReturn
  EndIf
  
  DropFolderPath = PathRequester("Select your folder with the installation files.", GetUserDirectory(#PB_Directory_Documents))
  DropFolderPath = RTrim(DropFolderPath, "\")
  
  If DropFolderPath
    DropFolder(DropFolderPath)
  EndIf
EndProcedure

Procedure MyInstallerFile(EventType)
  Protected InstallerFile_Selected.s = GetGadgetText(Combo_InstallerFile)
  
  If EventType = #PB_EventType_Change
    ;Debug InstallerFile_Selected
    
    Select GetExtensionPart(InstallerFile_Selected)
        
      Case "exe"
        DisableGadget(Button_MSI, #True)
        DisableGadget(Button_PEiD, #False)
        DisableGadget(Button_RunInstaller, #False)
        DisableGadget(Button_PSAppDeployToolkit, #False)
        
      Case "msi"
        DisableGadget(Button_MSI, #False)
        DisableGadget(Button_PEiD, #True)
        DisableGadget(Button_RunInstaller, #False)
        DisableGadget(Button_PSAppDeployToolkit, #False)
        
      Default
        DisableGadget(Button_MSI, #True)
        DisableGadget(Button_PEiD, #True)
        DisableGadget(Button_RunInstaller, #True)
        DisableGadget(Button_PSAppDeployToolkit, #False)
        
    EndSelect
  EndIf
EndProcedure

Procedure OpenPackageFolder(EventType)
  RunProgram("explorer.exe", PackagePath, "")
EndProcedure

Procedure OpenPSADT_Folder(EventType)
  RunProgram("explorer.exe", PSADT_Path, "")
EndProcedure

Procedure OpenSoftwareLogs_Folder(EventType)
  RunProgram("explorer.exe", "C:\Windows\Logs\Software", "")
EndProcedure

Procedure PSADT_StartPowerShellEditor(EventType)
  ShellExecute_(0, "RunAS", "powershell_ise.exe", Chr(34) + DropFolderPath + "\Deploy-Application.ps1" + Chr(34), "", #SW_SHOWNORMAL)
EndProcedure

Procedure PSADT_StartInstallation(EventType)
  ShellExecute_(0, "RunAS", Chr(34) + DropFolderPath + "\Deploy-Application.exe" + Chr(34), "-DeploymentType 'Install'", "", #SW_SHOWNORMAL)
EndProcedure

Procedure PSADT_StartUninstall(EventType)
  ShellExecute_(0, "RunAS", Chr(34) + DropFolderPath + "\Deploy-Application.exe" + Chr(34), "-DeploymentType 'Uninstall'", "", #SW_SHOWNORMAL)
EndProcedure

Procedure PSADT_StartRepair(EventType)
  ShellExecute_(0, "RunAS", Chr(34) + DropFolderPath + "\Deploy-Application.exe" + Chr(34), "-DeploymentType 'Repair'", "", #SW_SHOWNORMAL)
EndProcedure

Procedure PSADT_StartPowerShell(EventType)
  ShellExecute_(0, "RunAS", "powershell.exe", "", DropFolderPath, #SW_SHOWNORMAL)
EndProcedure

Procedure PSADT_StartHelp(EventType)
  ShellExecute_(0, "RunAS", "powershell.exe", "-ExecutionPolicy ByPass -File " + Chr(34) + DropFolderPath + "\AppDeployToolkit\AppDeployToolkitHelp.ps1" + Chr(34) + "", "", #SW_SHOWNORMAL)
EndProcedure

Procedure ListFilesRecursive(Dir.s, List Files.s())
  Protected D
  NewList Directories.s()
  
  If Right(Dir, 1) <> "\"
    Dir + "\"
  EndIf
  
  D = ExamineDirectory(#PB_Any, Dir, "")
  While NextDirectoryEntry(D)
    
    Select DirectoryEntryType(D)
      Case #PB_DirectoryEntry_File
        AddElement(Files())
        Files() = Dir + DirectoryEntryName(D)
      Case #PB_DirectoryEntry_Directory
        Select DirectoryEntryName(D)
          Case ".", ".."
            Continue
          Default
            AddElement(Directories())
            Directories() = Dir + DirectoryEntryName(D)
        EndSelect
    EndSelect
  Wend
  
  FinishDirectory(D)
  
  ForEach Directories()
    ListFilesRecursive(Directories(), Files())
  Next
  
EndProcedure

Procedure PSADT_GenerateExecutableList(EventType)
  Protected InstallationPath.s = PathRequester("Path to the installation folder", "C:\Program Files")
  Protected ExecutableList.s = ""
  
  If InstallationPath
    NewList F.s()
    ListFilesRecursive(InstallationPath, F())
    
    ForEach F()
      If GetExtensionPart(F()) = "exe"
        ExecutableList = ExecutableList + "," + GetFilePart(F(), #PB_FileSystem_NoExtension)
      EndIf
    Next
    
    ExecutableList = LTrim(ExecutableList, ",")
    InputRequester("Executable List", "Here is the result:", ExecutableList)
  EndIf

EndProcedure

Procedure RunInstallerFile(EventType)
  Protected InstallerFile.s = GetGadgetText(Combo_InstallerFile)
  RunProgram(DropFolderPath + "\" + InstallerFile, "", "")
EndProcedure

Procedure UpdateIntuneCommands()
  ; Example: https://www.google.com/search?q=vlc.exe+icon&tbm=isch
  Protected InstallerFile.s = GetGadgetText(Combo_InstallerFile)
  Protected Intune_InstallCommand.s, Intune_UninstallCommand.s, Intune_DetectionCommand.s
  Define Msi_ProductCode.s
  Define ReadMsi = #False
  
  Select GetExtensionPart(InstallerFile)
      
    Case "ps1"
      Intune_InstallCommand = "powershell.exe -ExecutionPolicy ByPass -File " + InstallerFile
      Intune_UninstallCommand = "powershell.exe -ExecutionPolicy ByPass -File " + InstallerFile
      Intune_DetectionCommand = "If (Test-Path " + Chr(34) + "C:\windows\system32\notepad2.exe" + Chr(34) + ") { Write-Output " + Chr(34) + "Notepad detected, exiting" + Chr(34) + " exit 0 } Else { exit 1 }"
      
    Case "cmd"
      Intune_InstallCommand = InstallerFile
      Intune_UninstallCommand = "<Uninstall Script>.cmd"
      Intune_DetectionCommand = "https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-add#step-4-detection-rules"
      
    Case "bat"
      Intune_InstallCommand = InstallerFile
      Intune_UninstallCommand = "<Uninstall Script>.bat"
      Intune_DetectionCommand = "https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-add#step-4-detection-rules"
      
    Case "exe"
      Intune_InstallCommand = InstallerFile + " /<Install Switch>"
      Intune_UninstallCommand = InstallerFile + " /<Uninstall Switch>"
      Intune_DetectionCommand = "https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-add#step-4-detection-rules"
      
    Case "msi"
      ReadMsi = #True
      Intune_InstallCommand = "(Please wait...)"
      Intune_UninstallCommand = "(Please wait...)"
      Intune_DetectionCommand = "(Please wait...)"
      
    Default
      ; Nothing
      
  EndSelect
  
  ; New: Support for PSAppDeployToolkit
  If InstallerFile = "Deploy-Application.exe" Or InstallerFile = "Deploy-Application.ps1"
      Intune_InstallCommand = "Deploy-Application.exe -DeploymentType 'Install' -DeployMode 'Silent'"
      Intune_UninstallCommand = "Deploy-Application.exe -DeploymentType 'Uninstall' -DeployMode 'Silent'"
      Intune_DetectionCommand = "https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-add#step-4-detection-rules"
  EndIf
  
  ; Read MSI
  If ReadMsi = #True
      ;Debug "Read MSI..."
      Msi_ProductCode = PowerShell_MsiProductCode(InstallerPath)
      Intune_InstallCommand = "msiexec /i " + Chr(34) + InstallerFile + Chr(34) + " /qn"
      Intune_UninstallCommand = "msiexec /x " + Chr(34) + Msi_ProductCode + Chr(34) + " /qn"
      Intune_DetectionCommand = Msi_ProductCode
  EndIf
  
  ; Update Gadgets
  SetGadgetText(Hyperlink_InstallCommand, Intune_InstallCommand)
  SetGadgetText(Hyperlink_UninstallCommand, Intune_UninstallCommand)
  SetGadgetText(Hyperlink_DetectionMethod, Intune_DetectionCommand)
EndProcedure

Procedure StartProcessIdentification(EventType)
  
  Protected InstallerFile.s, InstallerPath.s
  Protected Compiler, Output$, Exitcode = 1
  
  InstallerFile = GetGadgetText(Combo_InstallerFile)
  InstallerPath = DropFolderPath + "\" + InstallerFile
  
  ; Run Program
  RunProgram(PEiD, "-hard " + Chr(34) + InstallerPath + Chr(34), "", #PB_Program_Open)
EndProcedure

Procedure CreatePackage(EventType)
  
  Protected WrapperParameters$, Compiler, Output$, Exitcode = 1
  
  DropFolderPath = RTrim(DropFolderPath, "\")
  InstallerFile = GetGadgetText(Combo_InstallerFile)
  InstallerPath = DropFolderPath + "\" + InstallerFile
  PackagePath = GetParentDirectory(DropFolderPath) + "Package"
  
  ;Debug "Installer path is: " + InstallerPath
  ;Debug "Package path is: " + PackagePath
  
  If Trim(InstallerFile) = ""
    ProcedureReturn MessageRequester("Installer file", "Please select first the installer file in the dropdown.", #PB_MessageRequester_Ok | #PB_MessageRequester_Warning)
  EndIf
  
  ;Debug "Folder path is: " + DropFolderPath
  ;Debug "Installer path is: " + InstallerPath
  ;Debug "Package path will be: " + PackagePath
  
  ; Create main package folder if not exist
  CreateDirectory(PackagePath)
  
  ; Create sub-folder for packages
  Protected SubFolderChecked = GetGadgetState(Checkbox_CreateSubFolders)
  
  If SubFolderChecked
    Protected Date$, Time$
    Date$ = FormatDate("%yyyy.%mm.%dd", Date())
    PackagePath = PackagePath + "\" + Date$
    CreateDirectory(PackagePath)
  EndIf
  
  ; Check if *intunewin file exists in the package folder
  If ExamineDirectory(0, PackagePath, "*.intunewin")
    While NextDirectoryEntry(0)
      If DirectoryEntryType(0) = #PB_DirectoryEntry_File
        
        Protected SearchedPackageFile.s = GetFilePart(InstallerFile, #PB_FileSystem_NoExtension) + ".intunewin"
        
        ;Debug "Installer is: " + GetFilePart(InstallerFile)
        ;Debug "(Searched) Installer Package is: " + SearchedPackageFile
        ;Debug "Directory File is: " + DirectoryEntryName(0)
        
        If DirectoryEntryName(0) = SearchedPackageFile
          ProcedureReturn MessageRequester("Existing Intune file", "Please delete first the existing package to continue: " + DirectoryEntryName(0), #PB_MessageRequester_Ok | #PB_MessageRequester_Warning)
        EndIf
      EndIf
    Wend
    FinishDirectory(0)
  EndIf
  
  ; Check if IntuneWinAppUtil.exe exists in the current folder
  If FileSize(IntuneWinAppUtil) = -1
    ProcedureReturn MessageRequester("Win32 Content Prep Tool", IntuneWinAppUtil + " is missing. Please download it from: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool", #PB_MessageRequester_Ok | #PB_MessageRequester_Warning)
  EndIf

  ; Define parameters for IntuneWinAppUtil
  WrapperParameters$ = "-c %FolderPath% -s %InstallerPath% -o %PackagePath%"
  WrapperParameters$ = ReplaceString(WrapperParameters$, "%FolderPath%", Chr(34) + DropFolderPath + Chr(34))
  WrapperParameters$ = ReplaceString(WrapperParameters$, "%InstallerPath%", Chr(34) + InstallerPath + Chr(34))
  WrapperParameters$ = ReplaceString(WrapperParameters$, "%PackagePath%", Chr(34) + PackagePath + Chr(34))
  
  ; Disable button
  DisableGadget(Button_CreatePackage, #True)
  SetGadgetText(Button_CreatePackage, "Please wait...")
  
  ; Run IntuneWinAppUtil]
  Compiler = RunProgram(IntuneWinAppUtil, WrapperParameters$, "", #PB_Program_Open) ; #PB_Program_Open | #PB_Program_Read | #PB_Program_Hide
  Output$ = ""
  
  If Compiler
    While ProgramRunning(Compiler)
      If AvailableProgramOutput(Compiler)
        Output$ + ReadProgramString(Compiler) + Chr(13)
      EndIf
    Wend
    
    Exitcode = ProgramExitCode(Compiler)
    Output$ + Chr(13) + Chr(13)
    Output$ + "Exitcode: " + Str(Exitcode)
    
    CloseProgram(Compiler)
  EndIf
  
  ; Enable button
  DisableGadget(Button_CreatePackage, #False)
  
  ; Check Exitcode
  If Exitcode = 0
    UpdateIntuneCommands()
    HideGadget(Image_GreenCheck, #False)
    HideGadget(Hyperlink_PackageFolder, #False)
    HideGadget(Hyperlink_PSADT_Folder, #True)
    HideGadget(Button_DebugPSADT, #True)
    SetGadgetText(Button_CreatePackage, "Create Package")
    ResizeWindow(MainWindow, WindowX(MainWindow), WindowY(MainWindow), WindowWidth(MainWindow), 860)
    DisableGadget(Button_UploadIntune, #False)
    SetGadgetText(Button_UploadIntune, "Upload to Intune")
  Else
    MessageRequester("Output", Output$, #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
    SetGadgetText(Button_CreatePackage, "Re-try packaging")
    ResizeWindow(MainWindow, WindowX(MainWindow), WindowY(MainWindow), WindowWidth(MainWindow), 520)
    DisableGadget(Button_UploadIntune, #True)
    HideGadget(Hyperlink_PSADT_Folder, #True)
    HideGadget(Button_DebugPSADT, #True)
    SetGadgetText(Button_UploadIntune, "Upload to Intune")
  EndIf
  
EndProcedure

Procedure GeneratePSADTProject(Parameter)
  
  Protected OverridePath
  Protected InstallerFile_Extension.s, TemplateFile.s
  
  PSADT_Path = GetParentDirectory(DropFolderPath) + "PSADT\"
  InstallerFile = GetGadgetText(Combo_InstallerFile)
  InstallerFile_Extension = GetExtensionPart(InstallerFile)
  
  ; Define template file
  SetGadgetText(Text_Introduction, "Identifying installer file type...")
  Select InstallerFile_Extension
      
    Case "exe"
      TemplateFile = "EXE.ps1"
      
    Case "msi"
      TemplateFile = "MSI.ps1"
      
    Case "bat"
      TemplateFile = "Batch.ps1"
      
    Case "cmd"
      TemplateFile = "Batch.ps1"
      
    Case "ps1"
      TemplateFile = "PowerShell.ps1"
      
    Default
      ; By default is EXE
      TemplateFile = "EXE.ps1"
      
  EndSelect
  
  ; Check if PSADT folder is already exists
  SetGadgetText(Text_Introduction, "Checking destination folder...")
  If ExamineDirectory(#PB_Any, PSADT_Path, "*.*")
    OverridePath = MessageRequester("Error", "Do you want to overwrite the existing folder? " + Chr(13) + PSADT_Path, #PB_MessageRequester_YesNoCancel | #PB_MessageRequester_Error)
    
    If OverridePath = #PB_MessageRequester_No Or OverridePath = #PB_MessageRequester_Cancel
      ProcedureReturn #False
    EndIf
    
  Else
    CreateDirectory(PSADT_Path)
  EndIf
  
  ; Copy PSADT files
  SetGadgetText(Text_Introduction, "Copying PSADT toolkit to destination folder...")
  CopyDirectory(GetCurrentDirectory() + "ThirdParty\PSAppDeployToolkit", PSADT_Path, "", #PB_FileSystem_Recursive)
  
  ; Copy setup files
  SetGadgetText(Text_Introduction, "Copying installation files to destination folder...")
  CopyDirectory(DropFolderPath, PSADT_Path + "Files", "", #PB_FileSystem_Recursive)
  DeleteFile(PSADT_Path + "Deploy-Application.ps1")
  
  ; Copy deployment file
  ;CopyFile(GetCurrentDirectory() + "Templates\PSAppDeployToolkit\Custom.ps1", PSADT_Path + "Deploy-Application.ps1")
  
  ; Create empty deployment file
  SetGadgetText(Text_Introduction, "Creating new deployment file...")
  Protected NewDeploymentFile = CreateFile(#PB_Any, PSADT_Path + "Deploy-Application.ps1") 
  CloseFile(NewDeploymentFile)
  
  ; Update deployment file by MSI
  If InstallerFile_Extension = "msi"
    SetGadgetText(Text_Introduction, "Reading MSI informations - This takes longer...")
    Powershell_MsiDetails2()
  Else    
  
    If CreationWindowActive
      Details\Publisher = GetGadgetText(CW_String_Publisher)
      Details\Name = GetGadgetText(CW_String_Productname)
      Details\Version = GetGadgetText(CW_String_Version)
    Else
      Details\Publisher = InputRequester("Publisher", "Please enter the publisher name:", "")
      Details\Name = InputRequester("Product name", "Please enter the product name:", GetFilePart(InstallerFile, #PB_FileSystem_NoExtension))
      Details\Productcode = ""
      Details\Version = InputRequester("Product version", "Please enter the product version:", "")
    EndIf
  EndIf
  
  Protected FileIn, FileOut, sLine.s
  FileIn = ReadFile(#PB_Any, GetCurrentDirectory() + "Templates\PSAppDeployToolkit\" + TemplateFile, #PB_File_SharedRead)
  If FileIn
    FileOut = OpenFile(#PB_Any, PSADT_Path + "Deploy-Application.ps1")
    If FileOut
      ; Read each line from the input file, replace text, and write to the output file
      While Not Eof(FileIn)
        sLine = ReadString(FileIn)
        sLine = ReplaceString(sLine, "<ProductPublisher>", Details\Publisher)
        sLine = ReplaceString(sLine, "<ProductName>", Details\Name)
        sLine = ReplaceString(sLine, "<ProductVersion>", Details\Version)
        sLine = ReplaceString(sLine, "<ProductCode>", Details\Productcode)
        sLine = ReplaceString(sLine, "<ScriptDate>", FormatDate("%dd/%mm/%yyyy", Date()))
        sLine = ReplaceString(sLine, "<ScriptAuthor>", UserName())
        sLine = ReplaceString(sLine, "<InstallerFile>", InstallerFile)
        
        If CreationWindowActive
          sLine = ReplaceString(sLine, "<InstallationPath>", GetGadgetText(CW_String_InstallationPath))
          sLine = ReplaceString(sLine, "<InstallationParameter>", GetGadgetText(CW_String_InstallationParameter))
          sLine = ReplaceString(sLine, "<UninstallationPath>", GetGadgetText(CW_String_UninstallationPath))
          sLine = ReplaceString(sLine, "<UninstallationParameter>", GetGadgetText(CW_String_UninstallationParameter))
          
          If PSADT_CloseApps = ""
            PSADT_CloseApps = "iexplore"
          EndIf
          
          sLine = ReplaceString(sLine, "<CloseApps>", PSADT_CloseApps)
        EndIf
        
        WriteString(FileOut, sLine + #CRLF$)
      Wend
      CloseFile(FileOut)
    EndIf
    CloseFile(FileIn)
  EndIf
  
  ; Finalizing
  SetGadgetText(Text_Introduction, "Finished.")
  HideGadget(Hyperlink_PackageFolder, #True)
  HideGadget(Hyperlink_PSADT_Folder, #False)
  HideGadget(Button_DebugPSADT, #True)
  DisableGadget(Button_PSAppDeployToolkit, #False)
  DisableGadget(Button_CreatePackage, #False)
  DisableGadget(Combo_InstallerFile, #False)
  
  CloseCreationWindow(0)
  CreationWindowActive = #False
  
  ; Ask for Import PSADT
  Protected ImportPSADT = MessageRequester("Import", "Project was generated. Do you want to import it?", #PB_MessageRequester_Info | #PB_MessageRequester_YesNo)
  
  If ImportPSADT = #PB_MessageRequester_Yes
    DropFolderPath = RTrim(PSADT_Path, "\")
    DropFolder(DropFolderPath)
  EndIf

EndProcedure

Procedure CreatePSADTProject(EventType)

  Protected InstallerFile_Extension.s

  InstallerFile = GetGadgetText(Combo_InstallerFile)
  InstallerFile_Extension = GetExtensionPart(InstallerFile)
  
  DisableGadget(Button_PSAppDeployToolkit, #True)
  DisableGadget(Button_RunInstaller, #True)
  DisableGadget(Button_CreatePackage, #True)
  DisableGadget(Combo_InstallerFile, #True)
  
  ; Open Creation Window
  If InstallerFile_Extension <> "msi" And CreationWindowActive = #False
    Debug "Show-up creation window..."
    ShowCreationWindow(0)
    CreationWindowActive = #True
    
    ; Fill strings
    SetGadgetText(CW_String_Productname, GetFilePart(InstallerFile, #PB_FileSystem_NoExtension))
    SetGadgetText(CW_String_InstallationPath, InstallerFile)
    
    ; Read PSADT.setuppackager
    If OpenPreferences(DropFolderPath + "\PSADT.setuppackager")
      PreferenceGroup("General")
      SetGadgetText(CW_String_Publisher, ReadPreferenceString("Publisher", ""))
      SetGadgetText(CW_String_Productname, ReadPreferenceString("Productname", ""))
      SetGadgetText(CW_String_Version, ReadPreferenceString("Version", ""))
      SetGadgetText(CW_String_InstallationPath, ReadPreferenceString("InstallationPath", InstallerFile))
      SetGadgetText(CW_String_InstallationParameter, ReadPreferenceString("InstallationParameter", ""))
      SetGadgetText(CW_String_UninstallationPath, ReadPreferenceString("UninstallationPath", ""))
      SetGadgetText(CW_String_UninstallationParameter, ReadPreferenceString("UninstallationParameter", ""))
      PSADT_CloseApps = ReadPreferenceString("CloseApps", "iexplore")
      
      ClosePreferences()
    Else
      Debug "PSADT.setuppackager not found."
    EndIf
    
    ProcedureReturn 0
  EndIf
  
  CreateThread(@GeneratePSADTProject(), 0)
EndProcedure

Procedure CopyProductcodeToClipboard(EventType)
  Protected ClipboardText.s = GetGadgetItemText(DW_ListIcon_ProductIDS, 0)
  SetClipboardText(ClipboardText)
  HideGadget(Image_CopyProductcode_Check, #False)
EndProcedure

Procedure SearchIcon(EventType)
  ; Example: https://www.google.com/search?q=vlc.exe+icon&tbm=isch
  Protected Search.s = GetGadgetText(Combo_InstallerFile)
  
  Select GetExtensionPart(Search)
    Case "ps1"
      Search = "PowerShell"
      
    Case "cmd"
      Search = "Windows CMD script"
      
    Case "bat"
      Search = "Windows batch script"
      
    Default
      ; Nothing
      
  EndSelect
  
  RunProgram("https://www.google.com/search?q=" + EscapeString(Search) + "+icon&tbm=isch", "", "")
EndProcedure

Procedure SearchDescription(EventType)
  ; Example: https://www.google.com/search?q=vlc.exe+icon&tbm=isch
  Protected Search.s = GetGadgetText(Combo_InstallerFile)
  RunProgram("https://www.google.com/search?q=" + EscapeString(Search) + "+description", "", "")
EndProcedure

Procedure SearchSilentSwitch(EventType)
  ; Example: https://www.google.com/search?q=vlc.exe+icon&tbm=isch
  Protected Search.s = GetGadgetText(Combo_InstallerFile)
  RunProgram("https://www.google.com/search?q=" + EscapeString(Search) + "+silent+switch", "", "")
EndProcedure

Procedure Clipboard_InstallCommand(EventType)
  Protected ClipboardText.s = GetGadgetText(Hyperlink_InstallCommand)
  SetClipboardText(ClipboardText)
EndProcedure

Procedure Clipboard_UninstallCommand(EventType)
  Protected ClipboardText.s = GetGadgetText(Hyperlink_UninstallCommand)
  SetClipboardText(ClipboardText)
EndProcedure

Procedure Clipboard_DetectionCommand(EventType)
  Protected ClipboardText.s = GetGadgetText(Hyperlink_DetectionMethod)
  SetClipboardText(ClipboardText)
EndProcedure

Procedure BrowseInstallationPath(EventType)
  Protected File.s, InstallationPath.s
  File = OpenFileRequester("Select your installation file", DropFolderPath + "\" + InstallerFile, "All files (*.*)|*.*", 0)
  
  If File
    InstallationPath = LTrim(ReplaceString(File, DropFolderPath, ""), "\")
    SetGadgetText(CW_String_InstallationPath, InstallationPath)
  EndIf
  
EndProcedure

Procedure BrowseUninstallationPath(EventType)
  Protected File.s
  File = OpenFileRequester("Select your uninstallation file", DropFolderPath + "\" + InstallerFile, "All files (*.*)|*.*", 0)
  
  If File
    SetGadgetText(CW_String_UninstallationPath, File)
  EndIf
  
EndProcedure

;- Initalize Main Window
OpenMainWindow()
ResizeWindow(MainWindow, WindowX(MainWindow), WindowY(MainWindow), WindowWidth(MainWindow), 520) 
EnableGadgetDrop(Image_DropFolder, #PB_Drop_Files, #PB_Drag_Link)
HideGadget(Image_GreenCheck, #True)
HideGadget(Hyperlink_PackageFolder, #True)
HideGadget(Hyperlink_PSADT_Folder, #True)
CreateThread(@DownloadIntuneWinAppUtil(), 0)

;- Event Loop
Repeat
  Event = WaitWindowEvent()

  Select EventWindow()
      
    ; Main Window
    Case MainWindow
      If Event = #PB_Event_CloseWindow
        End
      ElseIf Event = #PB_Event_GadgetDrop
        DropFolderPath = EventDropFiles()
        
        If GetExtensionPart(DropFolderPath) = "" 
          DropFolder(DropFolderPath)
        EndIf
      Else
        MainWindow_Events(Event)
      EndIf
      
    ; About Window
    Case AboutWindow
      If Event = #PB_Event_CloseWindow
        CloseAboutWindow(0)
      Else
        AboutWindow_Events(Event)
      EndIf
      
    ; Details Window - MSI
    Case DetailsWindow_MSI
      If Event = #PB_Event_CloseWindow
        CloseDetailsWindow(0)
      Else
        DetailsWindow_MSI_Events(Event)
      EndIf
      
    ; Debugger Window - PSADT
    Case DebuggerWindow_PSADT
      If Event = #PB_Event_CloseWindow
        CloseDebuggerWindow(0)
      Else
        DebuggerWindow_PSADT_Events(Event)
      EndIf
      
     ; Debugger Window - PSADT
    Case CreationWindow_PSADT
      If Event = #PB_Event_CloseWindow
        CloseCreationWindow(0)
      Else
        CreationWindow_PSADT_Events(Event)
      EndIf
      
      ; Search Window - Installed Apps
    Case SearchWindow_InstalledApps
      If Event = #PB_Event_CloseWindow
        CloseSearchWindow(0)
      ElseIf Event = #PB_Event_Menu
          Select EventMenu()
            Case #SearchWindow_EnterPressed : SearchInstalledApp(#SearchWindow_EnterPressed)
          EndSelect
      Else
        SearchWindow_InstalledApps_Events(Event)
      EndIf
      
  EndSelect
  
Until Quit = #True

; IDE Options = PureBasic 6.11 LTS (Windows - x64)
; CursorPosition = 1192
; FirstLine = 173
; Folding = AAAAAAAAw
; EnableXP
; DPIAware