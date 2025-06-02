Attribute VB_Name = "Module1"
Option Explicit  ' Force variable declaration (�ѧ�Ѻ����С�ȵ���á�͹��ҹ)

'--- Declare SolidWorks API objects (��С�ȵ��������Ѻ��Ѻ SolidWorks API) ---
Public swApp As Object  ' SolidWorks application object (��ͺਡ��ͧ����� SolidWorks)
Public swModel As ModelDoc2  ' Active document (�͡��÷���Դ����㹢�й��)
Public swModelExt As IModelDocExtension  ' Document extension (��ǹ���¢ͧ�͡���)
Public cusPropMgr As CustomPropertyManager  ' Custom property manager (�Ѵ��äس���ѵԾ����)
Public configMgr As IConfigurationManager  ' Configuration manager (�Ѵ��ä͹�ԡ)
Public config As IConfiguration  ' Current configuration (�͹�ԡ�Ѩ�غѹ)

'--- Create instances of clsConfig (���ҧ��ͺਡ�� clsConfig ����Ѻ�红����� config) ---
Public preConf As New clsConfig
Public newConf As New clsConfig

Sub main()
    
    '********************************************************************
    Debug.Print "---------------------------------"
    Debug.Print "Step 01 : Start Sub Main"
    '********************************************************************
    
    '--- Show UserForm (�ʴ�������������ҹ) ---
    SaveForm.Show

    '********************************************************************
    Debug.Print "---------------------------------"
    Debug.Print "Step 99 : End Sub Main"
    '********************************************************************

End Sub

