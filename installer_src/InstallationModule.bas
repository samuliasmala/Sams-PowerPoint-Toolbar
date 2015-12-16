Attribute VB_Name = "InstallationModule"
Option Explicit

Sub InstallSamsToolkit()
    Dim toolbarName As String, sourceFile As String, destinationFile As String
    Dim fso As Object
    Dim regKeyBasePath As String, addinRegPath As String
    Dim addinsLocalDirectory As String, addInsFolder As String
    toolbarName = "Sam's toolkit"
    
    'Registry paths
    regKeyBasePath = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\" & Application.Version
    addinRegPath = regKeyBasePath & "\PowerPoint\AddIns\" & toolbarName
    
    'Check if toolbar file exists
    sourceFile = ActivePresentation.Path & "\" & toolbarName & ".ppam"
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(sourceFile) Then
        MsgBox sourceFile & " not found", vbInformation, "Not Found"
        Exit Sub
    End If

    'Find out the correct addin location
    addinsLocalDirectory = RegKeyRead(regKeyBasePath & "\Common\General\AddIns")
    addInsFolder = Environ("AppData") + "\Microsoft\" + addinsLocalDirectory
    
    'Save addin file
    destinationFile = addInsFolder & "\" & toolbarName & ".ppam"
    'Call ActivePresentation.SaveCopyAs(ppamFilename, ppSaveAsOpenXMLAddin)
    Call fso.CopyFile(sourceFile, destinationFile)

    Call RegKeySave(addinRegPath & "\AutoLoad", 1, "REG_DWORD")
    Call RegKeySave(addinRegPath & "\Path", destinationFile, "REG_SZ")
    
    MsgBox toolbarName & " installed succesfully! Please restart PowerPoint to get the new in the ribbon"
End Sub
