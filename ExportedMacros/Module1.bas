Attribute VB_Name = "Module1"
Dim swApp As Object
'Call use IModelDoc2 Interface
Dim swModel As ModelDoc2
'Call use IModelDocExtension Interface
Dim swModelExt As IModelDocExtension
'Call use ICustomPropertyManager Interface
Dim cusPropMgr As CustomPropertyManager
'Call use IConfigurationManager
Dim configMgr As IConfigurationManager
'Call use IConfiguration
Dim config As IConfiguration


'................................................................................................................................


'Parameters for SaveAs New Version
Dim Name As String
Dim Version As Integer
Dim Options As Integer
Dim ExportData As Object
Dim AdvancedSaveAsOptions As Object
Dim Errors As Integer
Dim Warnings As Integer
Dim value As Boolean

'................................................................................................................................





Sub main()


Debug.Print "Step 06 : Start Sub Main"
'................................................................................................................................

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc
Set swModelExt = swModel.Extension
'Set cusPropMgr = swModelExt.CustomPropertyManager("")

'Debug.Print "Step 07 : Set swApp ActiveDoc"
'................................................................................................................................

'Call Form
        
    SaveForm.Show

Debug.Print "Step 10 : Call Form"
'...............................................................................................................................

'Check Create Folder

Dim FoldName As String
Dim FoldExists As String

FoldName = LoadSetting("TextBox_ChangePath")
FoldExists = Dir(FoldName, vbDirectory)

Debug.Print "Folder Name : " & FoldName

If FoldExists = "" Then
        Debug.Print "The selected folder doesn't exist"
        MkDir (FoldName)
        Debug.Print "Completed Create the folder."
    Else
        Debug.Print "The selected folder exists"
    End If


Debug.Print "Step 09 : Check Create Drawing STEP Folder"
'................................................................................................................................
    
'Check File Name

Dim FileName As String
Dim LenName As Long
LenName = Len(swModel.GetPathName()) - InStrRev(swModel.GetPathName(), "\")

    If LoadSetting("CheckBox_ChangeName") = "True" Then
        FileName = LoadSetting("TextBox_ChangeName")
    Else
        FileName = Left(Right(swModel.GetPathName(), LenName), InStrRev(Right(swModel.GetPathName(), LenName), ".") - 1)
    End If
    
    Debug.Print "FileName = " & FileName
'................................................................................................................................

'SaveAs to .STEP .x_t

'........................................................................................
'stepFileName = FoldName & PartNo & ".STEP"
'paraFileName = FoldName & PartNo & ".x_t"
'SaveOPT = 4 + 2 + 1
'SaveResult = swModelExt.SaveAs(stepFileName, 0, SaveOPT, Nothing, 1, 1)
'Debug.Print " Save Result : " & SaveResult
'SaveResult = swModelExt.SaveAs(paraFileName, 0, SaveOPT, Nothing, 1, 1)
'Debug.Print " Save Result : " & SaveResult

'........................................................................................

'SaveAs to .dxf .dwg .pdf

Debug.Print "Folder Name : " & FoldName
Debug.Print "FileName = " & FileName

If LoadSetting("CheckBox_DWG") = "True" Then
    dwgFileName = FoldName & "\" & FileName & ".dwg"
    Debug.Print dwgFileName
    SaveOPT = 1 + 2
    SaveResult = swModelExt.SaveAs3(dwgFileName, 0, SaveOPT, Nothing, Nothing, 1, 1)
    Debug.Print " Save Result .dwg : " & SaveResult
End If

If LoadSetting("CheckBox_DXF") = "True" Then
    dxfFileName = FoldName & "\" & FileName & ".dxf"
    Debug.Print dxfFileName
    SaveOPT = 1 + 2
    SaveResult = swModelExt.SaveAs3(dxfFileName, 0, SaveOPT, Nothing, Nothing, 1, 1)
    Debug.Print " Save Result .dxf : " & SaveResult
End If

If LoadSetting("CheckBox_PDF") = "True" Then
    pdfFileName = FoldName & "\" & FileName & ".pdf"
    Debug.Print pdfFileName
    SaveOPT = 1 + 2
    SaveResult = swModelExt.SaveAs3(pdfFileName, 0, SaveOPT, Nothing, Nothing, 1, 1)
    Debug.Print " Save Result .pdf : " & SaveResult
End If

Debug.Print "Step 11 : Save As Files"
'...............................................................................................................................

Debug.Print "Step XX : End of Program"
'...............................................................................................................................

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



