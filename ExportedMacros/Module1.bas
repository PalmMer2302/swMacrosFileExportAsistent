Attribute VB_Name = "Module1"
Option Explicit  ' Force variable declaration (บังคับให้ประกาศตัวแปรก่อนใช้งาน)

'--- Declare SolidWorks API objects (ประกาศตัวแปรสำหรับใช้กับ SolidWorks API) ---
Public swApp As Object  ' SolidWorks application object (อ็อบเจกต์ของโปรแกรม SolidWorks)
Public swModel As ModelDoc2  ' Active document (เอกสารที่เปิดอยู่ในขณะนั้น)
Public swModelExt As IModelDocExtension  ' Document extension (ส่วนขยายของเอกสาร)
Public cusPropMgr As CustomPropertyManager  ' Custom property manager (จัดการคุณสมบัติพิเศษ)
Public configMgr As IConfigurationManager  ' Configuration manager (จัดการคอนฟิก)
Public config As IConfiguration  ' Current configuration (คอนฟิกปัจจุบัน)

'--- Create instances of clsConfig (สร้างอ็อบเจกต์ clsConfig สำหรับเก็บข้อมูล config) ---
Public preConf As New clsConfig
Public newConf As New clsConfig

Sub main()
    
    '********************************************************************
    Debug.Print "---------------------------------"
    Debug.Print "Step 01 : Start Sub Main"
    '********************************************************************
    
    '--- Show UserForm (แสดงฟอร์มให้ผู้ใช้งาน) ---
    SaveForm.Show

    '********************************************************************
    Debug.Print "---------------------------------"
    Debug.Print "Step 99 : End Sub Main"
    '********************************************************************

End Sub

