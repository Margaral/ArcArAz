Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms

<ComClass(tool_7_1_CreateDomain.ClassId, tool_7_1_CreateDomain.InterfaceId, tool_7_1_CreateDomain.EventsId), _
 ProgId("ArcArAz.tool_7_1_CreateDomain")> _
Public NotInheritable Class tool_7_1_CreateDomain
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "5f5a7b18-fcf9-42ca-9a26-001e964cb702"
    Public Const InterfaceId As String = "184bbfea-d2e6-4d26-a524-537658241ec3"
    Public Const EventsId As String = "b350763a-ecd2-4c8b-bc2b-8699ec107e97"
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
        ' TODO: Define values for the public properties
        MyBase.m_category = "ArcArAz-Domain"  'localizable text 
        MyBase.m_caption = "Create Domain"   'localizable text 
        MyBase.m_message = "Create a domain from existing entities"   'localizable text 
        MyBase.m_toolTip = "Create a domain from existing entities" 'localizable text 
        MyBase.m_name = "ArcArAz-Domain_CreateDomainTool"  'unique id, non-localizable (e.g. "MyCategory_ArcMapTool")

        'Try
        '    'TODO: change resource name if necessary
        '    Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
        '    MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        '    MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), Me.GetType().Name + ".cur")
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
        'TODO: Add tool_7_1_CreateDomain.OnClick implementation
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add tool_7_1_CreateDomain.OnMouseDown implementation
    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add tool_7_1_CreateDomain.OnMouseMove implementation
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add tool_7_1_CreateDomain.OnMouseUp implementation
    End Sub
End Class

