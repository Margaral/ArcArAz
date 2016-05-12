Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.SystemUI

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Text
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.AnalysisTools
Imports ESRI.ArcGIS.DataManagementTools
Imports ESRI.ArcGIS.GeoprocessingUI


<ComClass(cmd_2_6_Line2Polygon.ClassId, cmd_2_6_Line2Polygon.InterfaceId, cmd_2_6_Line2Polygon.EventsId), _
 ProgId("ArcArAz.cmd_2_6_Line2Polygon")> _
Public NotInheritable Class cmd_2_6_Line2Polygon
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4d21f6b7-188e-42c3-a189-5251775c7a82"
    Public Const InterfaceId As String = "7b8d8e18-8c0f-4dbc-857e-c5c6aa74cf7a"
    Public Const EventsId As String = "e5731786-2336-4b45-9f49-ef8cb522b4de"
#End Region

#Region "COM Registration Function(s)"
    <ComRegisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub RegisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryRegistration(registerType)

        'Add any COM registration code after the ArcGISCategoryRegistration() call

    End Sub

    <ComUnregisterFunction(), ComVisibleAttribute(False)> _
    Public Shared Sub UnregisterFunction(ByVal registerType As Type)
        ' Required for ArcGIS Component Category Registrar support
        ArcGISCategoryUnregistration(registerType)

        'Add any COM unregistration code after the ArcGISCategoryUnregistration() call

    End Sub

#Region "ArcGIS Component Category Registrar generated code"
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Register(regKey)

    End Sub
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommands.Unregister(regKey)

    End Sub

#End Region
#End Region


    Private m_application As IApplication

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        MyBase.m_category = "ArcArAz-EntityProp"  'localizable text 
        MyBase.m_caption = "Line to Polygon"   'localizable text 
        MyBase.m_message = "Converts from line to polygon"   'localizable text 
        MyBase.m_toolTip = "Converts from line to polygon" 'localizable text 
        MyBase.m_name = "ArcArAz-EntityProp_Line2Polygon"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

        'Try
        '    'TODO: change bitmap name if necessary
        '    Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
        '    MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        'Catch ex As Exception
        '    System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        'End Try


    End Sub


    Public Overrides Sub OnCreate(ByVal hook As Object)
        If Not hook Is Nothing Then
            m_application = CType(hook, IApplication)

            'Disable if it is not ArcMap
            If TypeOf hook Is IMxApplication Then
                MyBase.m_enabled = True
            Else
                MyBase.m_enabled = False
            End If
        End If

        ' TODO:  Add other initialization code
    End Sub

    Public Overrides Sub OnClick()
        'TODO: Add cmd_2_6_Line2Polygon.OnClick implementation

        Dim arcToolBoxExtension As IArcToolboxExtension = m_application.FindExtensionByName("ESRI ArcToolbox")

        If Not arcToolBoxExtension Is Nothing Then
            Dim arcToolBox As IArcToolbox = arcToolBoxExtension.ArcToolbox
            Dim gpTool = arcToolBox.GetToolbyNameString("FeatureToPolygon")
            If Not gpTool Is Nothing Then
                arcToolBox.InvokeTool(m_application.hWnd, gpTool, Nothing, False)
            End If
        End If

        'Dim pToolHelper As IGPToolCommandHelper2 = New GPToolCommandHelper

        ''Set the tool you want to invoke.
        'Dim toolboxPath = GetArcGISInstallDir() & "ArcToolbox\Toolboxes\Data Management Tools.tbx"
        'Try
        '    pToolHelper.SetToolByName(toolboxPath, "FeatureToPolygon")

        '    'Create the messages object to pass to the InvokeModal method.
        '    Dim msgs As IGPMessages
        '    msgs = New GPMessages

        '    'Invoke the tool.
        '    pToolHelper.InvokeModal(0, Nothing, True, msgs)
        '    m_application.CurrentTool = Nothing
        'Catch err As Exception
        '    MsgBox("The tool instalation folder is not the default." & vbCrLf & _
        '            "Please, access this tool throught ArcToolBox: " & vbCrLf & _
        '            "ArcToolBox / Data Management Tools / Features / Feature To Polygon.")
        'End Try

    End Sub

    '    Public Function GetArcGISInstallDir() As String
    '        On Error GoTo Err
    '        Dim WScr As Object
    '        WScr = CreateObject("WScript.Shell")
    '        Dim sDir As String
    '        sDir = WScr.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\ESRI\ArcGIS\InstallDir")
    '        GetArcGISInstallDir = sDir
    '        If Trim(sDir) = vbNullString Then
    '            MsgBox("ERROR: ARCGIS-1" & vbCrLf & "Unable to retrieve ArcGIS information on this machine." & vbCrLf & "ArcGIS installation on this machine seems to be corrupt." & vbCrLf & "Please contact your system administrator.")
    '        End If
    '        GoTo Cleanup
    'Err:
    '        GetArcGISInstallDir = ""
    '        If Err.Number = 53 Then
    '            MsgBox("ERROR: ARCGIS-2" & vbCrLf & "ArcGIS is not installed on this machine." & vbCrLf & "Please install ArcGIS prior to run this application.")
    '        Else
    '            MsgBox("ERROR: " & Err.Number & vbCrLf & Err.Description)
    '        End If
    'Cleanup:
    '        WScr = Nothing
    '    End Function
End Class



