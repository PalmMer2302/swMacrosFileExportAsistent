VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveForm 
   Caption         =   "Save As"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9270.001
   OleObjectBlob   =   "SaveForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BrowseButton_Click()
     ' Setting Solidworks variable to Solidworks application
  Set swApp = Application.SldWorks
  
  ' Browse and get the Selected file name
  SaveAsPath = BrowseForFolder()
  ' Show the selected file's full path in text box
  TextBox_ChangePath.Text = SaveAsPath
  
End Sub

Private Sub Button_Cancel_Click()

    Call SaveSetting(False, "", False, "", False, False, False)

    Unload Me

End Sub

Private Sub Button_Save_Click()

    
    Call SaveSetting(CheckBox_ChangeName.value, TextBox_ChangeName.Text, CheckBox_ChangePath.value, TextBox_ChangePath.Text, CheckBox_DWG.value, CheckBox_DXF.value, CheckBox_PDF.value)

    
    Me.Hide
    

End Sub



Private Sub UserForm_Initialize()

    'GetName
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Dim LenName As Long
    LenName = Len(swModel.GetPathName()) - InStrRev(swModel.GetPathName(), "\")
    
    Dim FileName As String
    FileName = Left(Right(swModel.GetPathName(), LenName), InStrRev(Right(swModel.GetPathName(), LenName), ".") - 1)
    
    'GetPathName
    'CheckBox_ChangeName
    'TextBox_ChangeName
    'CheckBox_ChangePath
    'TextBox_ChangePath
    'CheckBox_DWG
    'CheckBox_DXF
    'CheckBox_PDF
    
    Dim fPathName As String
    fPathName = Environ("USERPROFILE") & "\Documents\SW_Macro_RPT\Drawing_Export\Setting_config.txt"
    
    If Dir(fPathName) = "" Then
        Debug.Print "The Setting File doesn't exist."
        
        Label_CurrentName.Caption = FileName
        CheckBox_ChangeName.value = False
        TextBox_ChangeName.Text = ""

    
        Label_CurrentPath.Caption = Left(swModel.GetPathName(), InStrRev(swModel.GetPathName(), "\"))
        CheckBox_ChangePath.value = False
        TextBox_ChangePath.Text = ""
    
        CheckBox_DWG.value = False
        CheckBox_DWG.Enabled = True
        CheckBox_DWG.Locked = False
        CheckBox_DXF.value = False
        CheckBox_DXF.Enabled = True
        CheckBox_DXF.Locked = False
        CheckBox_PDF.value = False
        CheckBox_PDF.Enabled = True
        CheckBox_PDF.Locked = False


    Else
        Debug.Print "The Setting File exists."
            
        Label_CurrentName.Caption = FileName
        CheckBox_ChangeName.value = LoadSetting("CheckBox_ChangeName")
        TextBox_ChangeName.Text = LoadSetting("TextBox_ChangeName")

    
        Label_CurrentPath.Caption = Left(swModel.GetPathName(), InStrRev(swModel.GetPathName(), "\"))
        CheckBox_ChangePath.value = LoadSetting("CheckBox_ChangePath")
        TextBox_ChangePath.Text = LoadSetting("TextBox_ChangePath")
    
        CheckBox_DWG.value = LoadSetting("CheckBox_DWG")
        CheckBox_DWG.Enabled = True
        CheckBox_DWG.Locked = False
        CheckBox_DXF.value = LoadSetting("CheckBox_DXF")
        CheckBox_DXF.Enabled = True
        CheckBox_DXF.Locked = False
        CheckBox_PDF.value = LoadSetting("CheckBox_PDF")
        CheckBox_PDF.Enabled = True
        CheckBox_PDF.Locked = False
        
    End If
    
End Sub


'**********************
'Copyright(C) 2023 Xarial Pty Limited
'Reference: https://www.codestack.net/visual-basic/algorithms/fso/browse-folder/
'License: https://www.codestack.net/license/
'**********************

Function BrowseForFolder() As String
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level

    Dim ShellApp As Object

     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a folder", 0)

     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0

     'Destroy the Shell Application
    Set ShellApp = Nothing

     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select

    Exit Function

Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = False
End Function

Public Sub SaveSetting(CBname As Boolean, TBname As String, CBpath As Boolean, TBpath As String, CBdwg As Boolean, CBdxf As Boolean, CBpdf As Boolean)

    Debug.Print "Run Func SaveSetting"
    
    'Setup Location Setting File Path
    Dim fPath1 As String
    Dim fPath2 As String
    Dim fName As String
    fPath1 = Environ("USERPROFILE") & "\Documents\SW_Macro_RPT\"
    fPath2 = Environ("USERPROFILE") & "\Documents\SW_Macro_RPT\Drawing_Export\"
    fName = "Setting_config.txt"
    Debug.Print "File 1st Level Path is : " & fPath1
    Debug.Print "File Full Path is : " & fPath2
    Debug.Print "File Name is : " & fName
    
    'Check Exists the Folder
    
    If Dir(fPath2, vbDirectory) = "" Then
        Debug.Print "The folder doesn't exist."
        MkDir (fPath1)
        MkDir (fPath2)
        Debug.Print "Completed Create the folder."
    Else
        Debug.Print "The folder exists."
    End If

    'Create the file
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(fPath2 & fName)
    oFile.WriteLine "CheckBox_ChangeName_" & CBname & "/end"
    oFile.WriteLine "TextBox_ChangeName_" & TBname & "/end"
    oFile.WriteLine "CheckBox_ChangePath_" & CBpath & "/end"
    oFile.WriteLine "TextBox_ChangePath_" & TBpath & "/end"
    oFile.WriteLine "CheckBox_DWG_" & CBdwg & "/end"
    oFile.WriteLine "CheckBox_DXF_" & CBdxf & "/end"
    oFile.WriteLine "CheckBox_PDF_" & CBpdf & "/end"
    oFile.Close
    Debug.Print "Completed Create the Text File."
    Set fso = Nothing
    Set oFile = Nothing
    
    Debug.Print "End Func SaveSetting"

End Sub

Function LoadSetting(LoadStr As String) As String

Debug.Print "Run Func LoadSetting"

    Dim fPath2 As String
    Dim fName As String
    fPath2 = Environ("USERPROFILE") & "\Documents\SW_Macro_RPT\Drawing_Export\"
    fName = "Setting_config.txt"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.OpenTextFile(fPath2 & fName)
    
    Dim TextStr As String
    TextStr = oFile.ReadAll
    
    oFile.Close
    
    Debug.Print "TextStr = "
    Debug.Print TextStr
    
    
    'ReadTextStr
    Dim sStr As String
    sStr = LoadStr & "_"
    
    Debug.Print "sStr = " & sStr
    
    StartStrPos = InStr(TextStr, sStr)
    Debug.Print "StartStrPos = " & StartStrPos
    
    EndStrPos = InStr(TextStr, sStr) + Len(sStr)
    Debug.Print "EndStrPos = " & EndStrPos
    
    EndVeluePos = InStr(EndStrPos, TextStr, "/end")
    Debug.Print "EndVeluePos = " & EndVeluePos
    
    ValueStr = Mid(TextStr, EndStrPos, EndVeluePos - EndStrPos)
    Debug.Print "string_to_search = " & TextStr
    Debug.Print "start_position = " & EndStrPos
    Debug.Print "number_of_characters = " & EndVeluePos - EndStrPos
    Debug.Print "ValueStr = " & ValueStr
    
    LoadSetting = ValueStr
    

Debug.Print "End Func LoadSetting"
Debug.Print " "
Debug.Print "..................................................."
End Function


