Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.GeoprocessingUI

<ComClass(cmd3_1DEMfromAutoCAD.ClassId, cmd3_1DEMfromAutoCAD.InterfaceId, cmd3_1DEMfromAutoCAD.EventsId), _
 ProgId("ArcArAz.cmd3_1DEMfromAutoCAD")> _
Public NotInheritable Class cmd3_1DEMfromAutoCAD
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3b24d046-63ec-4aeb-a7a4-0e89ee6b4725"
    Public Const InterfaceId As String = "c030e9f5-de25-4888-aa3f-7ef210e4da8d"
    Public Const EventsId As String = "679e311d-342e-47e4-9e44-183d6a19d267"
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
        MyBase.m_category = "ArcArAz-SpatialFields"  'localizable text 
        MyBase.m_caption = "DEM from AutoCAD"   'localizable text 
        MyBase.m_message = "Create a raster or TIN representing topografy"   'localizable text 
        MyBase.m_toolTip = "Create a raster or TIN representing topografy" 'localizable text 
        MyBase.m_name = "ArcArAz-SpatialFields_DEMFromAutoCADCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd3_1DEMfromAutoCAD.OnClick implementation

        Dim arcToolBoxExtension As IArcToolboxExtension = m_application.FindExtensionByName("ESRI ArcToolbox")

        If Not arcToolBoxExtension Is Nothing Then
            Dim arcToolBox As IArcToolbox = arcToolBoxExtension.ArcToolbox
            Dim gpTool = arcToolBox.GetToolbyNameString("DEMfromCAD")
            If Not gpTool Is Nothing Then
                arcToolBox.InvokeTool(m_application.hWnd, gpTool, Nothing, False)
            End If
        End If

    End Sub
End Class



