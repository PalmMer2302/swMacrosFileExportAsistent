VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveForm 
   Caption         =   "Save As"
   ClientHeight    =   6486
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9270.001
   OleObjectBlob   =   "SaveForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  ' Force variable declaration (บังคับให้ประกาศตัวแปรก่อนใช้งาน)

'--- Declare variables for internal use ---
Private FileName As String        ' Current file name (ชื่อไฟล์ปัจจุบัน)
Private modelPath As String       ' Full path of current document (เส้นทางของไฟล์ SolidWorks)
Private fPathName As String       ' Config file path (พาธของไฟล์ config)

'=========================================
' Browse Button Click - เปิดเลือกโฟลเดอร์
'=========================================
Private Sub BrowseButton_Click()
     ' Setting Solidworks variable to Solidworks application
  Set swApp = Application.SldWorks
  Dim SaveAsPath As String
  ' Browse and get the Selected file name
  SaveAsPath = BrowseForFolder()
  ' Show the selected file's full path in text box
  TextBox_ChangePath.Text = SaveAsPath
End Sub

'=========================================
' Cancel Button Click - ปิดฟอร์ม
'=========================================
Private Sub Button_Cancel_Click()
    Debug.Print "-->Cancel"
    Unload Me
End Sub

'=========================================
' Load previous config button
'=========================================
Private Sub Button_PreConf_Click()
    Debug.Print "-->Load Pre Config"
    LoadValuesFromConfig
End Sub

'=========================================
' Save Button Click - บันทึกไฟล์หลายประเภท
'=========================================
Private Sub Button_Save_Click()
    Debug.Print "-->Save"
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swModelExt = swModel.Extension

    ' Handle change file name
    newConf.cbChangeName = CheckBox_ChangeName.Value
    If newConf.cbChangeName Then
        If TextBox_ChangeName.Text <> "" Then
            newConf.tbName = TextBox_ChangeName.Text
        Else
            MsgBox "Invalid file name." & vbCrLf & _
            "(ไม่พบชื่อไฟล์เอกสาร)", _
            vbExclamation, "Document Not Found"
            Exit Sub
        End If
    Else
        newConf.tbName = FileName
    End If

    ' Handle change path
    newConf.cbPath = CheckBox_ChangePath.Value
    If newConf.cbPath Then
        If TextBox_ChangePath.Text <> "" Then
            newConf.tbPath = TextBox_ChangePath.Text
        Else
        MsgBox "Invalid file path." & vbCrLf & _
            "(ไม่พบที่อยู่ไฟล์เอกสาร)", _
            vbExclamation, "Document Not Found"
            Exit Sub
        End If
    Else
        newConf.tbPath = Left(modelPath, InStrRev(modelPath, "\"))
    End If

    ' Export type selections
    newConf.cbDWG = CheckBox_DWG.Value
    newConf.cbDXF = CheckBox_DXF.Value
    newConf.cbPDF = CheckBox_PDF.Value

    ' Save to config
    newConf.saveConf newConf

    ' Check if folder exists
    If Dir(newConf.tbPath, vbDirectory) = "" Then MkDir (newConf.tbPath)

    ' Save files
    If newConf.cbDWG Then ExportFile newConf.tbName, newConf.tbPath, "DWG", "dwg"
    If newConf.cbDXF Then ExportFile newConf.tbName, newConf.tbPath, "DXF", "dxf"
    If newConf.cbPDF Then ExportFile newConf.tbName, newConf.tbPath, "PDF", "pdf"

    Me.Hide
End Sub

'=========================================
' UserForm Initialization - โหลดค่าเมื่อเปิดฟอร์ม
'=========================================
Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "No active SolidWorks document found." & vbCrLf & _
               "(ไม่พบไฟล์เอกสารที่กำลังใช้งานอยู่ในโปรแกรม SolidWorks)", _
               vbExclamation, "Document Not Found"
        Unload Me: End
    End If

    modelPath = swModel.GetPathName()
    If modelPath = "" Then
        MsgBox "This document has not been saved. Please save it first." & vbCrLf & _
               "(เอกสารนี้ยังไม่ได้รับการบันทึก กรุณาบันทึกก่อนดำเนินการต่อ)", _
               vbExclamation, "Save Required"
        Unload Me: End
    End If

    FileName = Mid(modelPath, InStrRev(modelPath, "\") + 1)
    FileName = Left(FileName, InStrRev(FileName, ".") - 1)

    fPathName = Environ("USERPROFILE") & "\Documents\SW_Macro_RPT\Drawing_Export\Setting_config.txt"

    If Dir(fPathName) = "" Then
        LoadDefaultValues
    Else
        LoadValuesFromConfig
    End If
    Exit Sub

ErrHandler:
    MsgBox "Unexpected error: " & Err.Description, vbCritical, "Macro Error"
    Unload Me
End Sub

'=========================================
' Load Default Values
'=========================================
Private Sub LoadDefaultValues()
    Debug.Print "-->LoadDefaultValues"
    Label_CurrentName.Caption = FileName
    CheckBox_ChangeName.Value = False
    TextBox_ChangeName.Text = ""

    Label_CurrentPath.Caption = Left(modelPath, InStrRev(modelPath, "\"))
    CheckBox_ChangePath.Value = False
    TextBox_ChangePath.Text = ""

    CheckBox_DWG.Value = False
    CheckBox_DXF.Value = False
    CheckBox_PDF.Value = False
End Sub

'=========================================
' Load Values from config
'=========================================
Private Sub LoadValuesFromConfig()
    Debug.Print "-->LoadValuesFromConfig"
    preConf.loadConf preConf

    Label_CurrentName.Caption = FileName
    CheckBox_ChangeName.Value = preConf.cbChangeName
    TextBox_ChangeName.Text = preConf.tbName

    Label_CurrentPath.Caption = Left(modelPath, InStrRev(modelPath, "\"))
    CheckBox_ChangePath.Value = preConf.cbPath
    TextBox_ChangePath.Text = preConf.tbPath

    CheckBox_DWG.Value = preConf.cbDWG
    CheckBox_DXF.Value = preConf.cbDXF
    CheckBox_PDF.Value = preConf.cbPDF
End Sub

'=========================================
' Export file to chosen format
'=========================================
Private Sub ExportFile(NewFileName As String, NewFolderPath As String, FileType As String, Extension As String)
    Dim saveResult As Boolean
    Dim errors As Long, warnings As Long
    Dim ExportPath As String
    ExportPath = NewFolderPath & "\" & NewFileName & "." & Extension

    ' Confirm before saving
    If MsgBox("Do you want to save the selected file(s)?" & vbCrLf & _
          "(คุณต้องการบันทึกไฟล์ที่เลือกไว้หรือไม่?)" & vbCrLf & _
          "File Name: " & NewFileName & vbCrLf & _
          "Folder Path: " & NewFolderPath, _
          vbYesNo + vbQuestion, "Confirm Save") = vbNo Then
        Exit Sub
    End If
    
    If Dir(ExportPath) <> "" Then
        If MsgBox("File already exists: " & ExportPath & vbCrLf & _
                  "Overwrite it?" & vbCrLf & _
                  "(ไฟล์นี้มีอยู่แล้ว ต้องการเขียนทับหรือไม่?)", _
                  vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    saveResult = swModelExt.SaveAs3(ExportPath, 0, 1 + 2, Nothing, Nothing, errors, warnings)

    If saveResult = False Or errors <> 0 Then
        MsgBox "Error saving " & Extension & " file: " & ExportPath & vbCrLf & _
               "Error code: " & errors & vbCrLf & _
               "(เกิดข้อผิดพลาดในการบันทึกไฟล์)", vbCritical
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
    
    If Not ShellApp Is Nothing Then
        On Error Resume Next
        BrowseForFolder = ShellApp.Self.Path
        On Error GoTo 0
    Else
        BrowseForFolder = ""
    End If
    
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

