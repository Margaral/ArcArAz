Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI

<ComClass(cmd_8_2_RasterTINDescensos.ClassId, cmd_8_2_RasterTINDescensos.InterfaceId, cmd_8_2_RasterTINDescensos.EventsId), _
 ProgId("ArcArAz.cmd_8_2_RasterTINDescensos")> _
Public NotInheritable Class cmd_8_2_RasterTINDescensos
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "25ea907a-b137-4302-8afa-1a63a7e727ea"
    Public Const InterfaceId As String = "fb4844dc-098b-438e-929d-1297fbaa76a5"
    Public Const EventsId As String = "7f79d5e9-c9f6-4b8f-a812-47137f272658"
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
        MyBase.m_category = "ArcArAz-Output"  'localizable text 
        MyBase.m_caption = "Raster / TIN with Descensos"   'localizable text 
        MyBase.m_message = "Create a raster or TIN with Descensos"   'localizable text 
        MyBase.m_toolTip = "Create a raster or TIN with Descensos" 'localizable text 
        MyBase.m_name = "ArcArAz-Output_RasterTINDescensos"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_8_2_RasterTINDescensos.OnClick implementation
    End Sub
End Class



