'Option Infer On
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.DataManagementTools
Imports ESRI.ArcGIS.CartographyTools
Imports ESRI.ArcGIS.GeoprocessingUI
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.esriSystem


<ComClass(cmd_9_1_SimplifyGraphic.ClassId, cmd_9_1_SimplifyGraphic.InterfaceId, cmd_9_1_SimplifyGraphic.EventsId), _
 ProgId("ArcArAz.cmd_9_1_SimplifyGraphic")> _
Public NotInheritable Class cmd_9_1_SimplifyGraphic
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "76b5826f-56ab-4ec5-97d9-6bf80af1a35d"
    Public Const InterfaceId As String = "8b16d130-7bfc-49c1-a2ab-8eb20e062d9a"
    Public Const EventsId As String = "ade0d3c8-6692-438b-b87c-35e9e707f1e2"
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
        MyBase.m_category = "ArcArAz-TemporalSeries"  'localizable text 
        MyBase.m_caption = "Simplify Graphic"   'localizable text 
        MyBase.m_message = "Add Excel file to TOC and select it"   'localizable text 
        MyBase.m_toolTip = "Add Excel file to TOC and select it" 'localizable text 
        MyBase.m_name = "ArcArAz-TemporalSeries_SimplifyGraphicz"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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

        Dim arcToolBoxExtension As IArcToolboxExtension = m_application.FindExtensionByName("ESRI ArcToolbox")

        If Not arcToolBoxExtension Is Nothing Then
            Dim arcToolBox As IArcToolbox = arcToolBoxExtension.ArcToolbox
            Dim gpTool = arcToolBox.GetToolbyNameString("ArcArAz1")
            If Not gpTool Is Nothing Then
                arcToolBox.InvokeTool(m_application.hWnd, gpTool, Nothing, False)
            End If
        End If

    End Sub

End Class



