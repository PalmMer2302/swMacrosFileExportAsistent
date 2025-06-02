VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************
' Class: clsConfig
' This class holds configuration settings and provides Save/Load methods
' คลาสนี้ใช้เก็บค่าการตั้งค่าและสามารถบันทึก/โหลดค่าจากไฟล์ได้
'*********************************************************************
Option Explicit  ' Force explicit declaration of all variables (บังคับให้ประกาศตัวแปรทุกตัว)

'==================
' Properties (คุณสมบัติที่ใช้เก็บค่าจากฟอร์ม)
'==================
Public cbChangeName As Boolean  ' Checkbox to enable filename change (เลือกเพื่อเปลี่ยนชื่อไฟล์)
Public tbName As String         ' Textbox value for new name (ชื่อไฟล์ใหม่)

Public cbPath As Boolean        ' Checkbox to enable custom path (เลือกเพื่อใช้ path ที่ผู้ใช้กำหนดเอง)
Public tbPath As String         ' Custom path input (เส้นทางกำหนดเอง)

Public cbDWG As Boolean         ' Export DWG (เลือกเพื่อ export DWG)
Public cbDXF As Boolean         ' Export DXF (เลือกเพื่อ export DXF)
Public cbPDF As Boolean         ' Export PDF (เลือกเพื่อ export PDF)

'==================
' Save configuration to file (บันทึกค่าการตั้งค่าลงในไฟล์)
'==================
Public Sub saveConf(ByRef conf As clsConfig)

    Debug.Print "---------------------------------"
    Debug.Print "--/Run clsConfig.save"

    Dim fPath1 As String, fPath2 As String, fName As String
    fPath1 = Environ("USERPROFILE") & "\Documents\SW_Macro_RPT\"          ' Base path
    fPath2 = fPath1 & "Drawing_Export\"                                    ' Subfolder
    fName = "Setting_config.txt"                                           ' Filename

    ' Create folders if they do not exist (สร้างโฟลเดอร์ถ้ายังไม่มี)
    If Dir(fPath1, vbDirectory) = "" Then MkDir fPath1
    If Dir(fPath2, vbDirectory) = "" Then MkDir fPath2

    ' Create file and write data
    Dim fso As Object, oFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(fPath2 & fName)

    ' Write each property as [Tag]_[Value]/end
    oFile.WriteLine "CheckBox_ChangeName_" & conf.cbChangeName & "/end"
    oFile.WriteLine "TextBox_ChangeName_" & conf.tbName & "/end"
    oFile.WriteLine "CheckBox_ChangePath_" & conf.cbPath & "/end"
    oFile.WriteLine "TextBox_ChangePath_" & conf.tbPath & "/end"
    oFile.WriteLine "CheckBox_DWG_" & conf.cbDWG & "/end"
    oFile.WriteLine "CheckBox_DXF_" & conf.cbDXF & "/end"
    oFile.WriteLine "CheckBox_PDF_" & conf.cbPDF & "/end"

    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

    Debug.Print "Completed writing settings to file."
    Debug.Print "End clsConfig.save"
    Debug.Print "---------------------------------"

End Sub

'==================
' Load configuration from file (โหลดค่าการตั้งค่าจากไฟล์)
'==================
Public Sub loadConf(ByRef conf As clsConfig)

    Debug.Print "---------------------------------"
    Debug.Print "--/Run clsConfig.load"

    Dim fPath2 As String, fName As String
    fPath2 = Environ("USERPROFILE") & "\Documents\SW_Macro_RPT\Drawing_Export\"
    fName = "Setting_config.txt"

    ' File system object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if file exists (ตรวจสอบว่าไฟล์ config มีอยู่หรือไม่)
    If Not fso.FileExists(fPath2 & fName) Then
        MsgBox _
        "Configuration file not found. Cannot load settings." & vbCrLf & _
        "(ไม่พบไฟล์การตั้งค่า จึงไม่สามารถโหลดค่าต่างๆ ได้)", _
        vbExclamation, "Missing File"
        Exit Sub
    End If

    ' Read full text from file
    Dim oFile As Object, TextStr As String
    Set oFile = fso.OpenTextFile(fPath2 & fName)
    TextStr = oFile.ReadAll
    oFile.Close

    ' Mapping table: tag in file to class property
    Dim prop_arr(6, 1) As String
    prop_arr(0, 0) = "CheckBox_ChangeName":  prop_arr(0, 1) = "cbChangeName"
    prop_arr(1, 0) = "TextBox_ChangeName":   prop_arr(1, 1) = "tbName"
    prop_arr(2, 0) = "CheckBox_ChangePath":  prop_arr(2, 1) = "cbPath"
    prop_arr(3, 0) = "TextBox_ChangePath":   prop_arr(3, 1) = "tbPath"
    prop_arr(4, 0) = "CheckBox_DWG":         prop_arr(4, 1) = "cbDWG"
    prop_arr(5, 0) = "CheckBox_DXF":         prop_arr(5, 1) = "cbDXF"
    prop_arr(6, 0) = "CheckBox_PDF":         prop_arr(6, 1) = "cbPDF"

    ' Extract and assign each value (ดึงค่าจากไฟล์และใส่ให้ property)
    Dim X As Integer
    Dim StartStrPos As Long, EndStrPos As Long, EndValuePos As Long
    Dim ValueStr As String, sStr As String

    For X = 0 To 6
        sStr = prop_arr(X, 0) & "_"
        StartStrPos = InStr(TextStr, sStr)

        If StartStrPos > 0 Then
            EndStrPos = StartStrPos + Len(sStr)
            EndValuePos = InStr(EndStrPos, TextStr, "/end")
            ValueStr = Mid(TextStr, EndStrPos, EndValuePos - EndStrPos)

            Debug.Print prop_arr(X, 1) & " = " & ValueStr

            ' Assign value dynamically (กำหนดค่าด้วย CallByName)
            CallByName conf, prop_arr(X, 1), VbLet, CBoolOrString(ValueStr)
        End If
    Next X

    Debug.Print "End clsConfig.load"
    Debug.Print "---------------------------------"

End Sub

'==================
' Helper: Convert "True"/"False" to Boolean, else keep as string
' ตัวช่วย: แปลงข้อความ "True"/"False" เป็น Boolean ถ้าไม่ใช่ก็คืนค่าข้อความเดิม
'==================
Private Function CBoolOrString(ByVal strValue As String) As Variant
    If LCase(strValue) = "true" Or LCase(strValue) = "false" Then
        CBoolOrString = CBool(strValue)
    Else
        CBoolOrString = strValue
    End If
End Function


