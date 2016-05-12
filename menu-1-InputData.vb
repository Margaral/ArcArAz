Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(menu1InputData.ClassId, menu1InputData.InterfaceId, menu1InputData.EventsId), _
 ProgId("ArcArAz.InputData")> _
Public NotInheritable Class menu1InputData
    Inherits BaseMenu

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
    ''' <summary>
    ''' Required method for ArcGIS Component Category registration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryRegistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Register(regKey)

    End Sub
    ''' <summary>
    ''' Required method for ArcGIS Component Category unregistration -
    ''' Do not modify the contents of this method with the code editor.
    ''' </summary>
    Private Shared Sub ArcGISCategoryUnregistration(ByVal registerType As Type)
        Dim regKey As String = String.Format("HKEY_CLASSES_ROOT\CLSID\{{{0}}}", registerType.GUID)
        MxCommandBars.Unregister(regKey)

    End Sub

#End Region
#End Region

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3c4884c3-64ff-4946-9284-c75cd1044df0"
    Public Const InterfaceId As String = "7cb82769-43bc-47ae-9f02-9ba01d9346c9"
    Public Const EventsId As String = "b0bc8f63-3f83-4abb-a69b-8fae0d5f9f0e"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        '
        'TODO: Define your menu here by adding items
        '
        'AddItem("esriArcMapUI.ZoomInFixedCommand")
        'BeginGroup() 'Separator
        AddItem("{54ce5f8c-3bf5-4dc1-bfa8-3c5bfce0df0e}", 1) 'Set project and workfolder command
        AddItem("{02e86e67-159b-40c3-9de9-7f92e4ab1e2f}", 1) 'Set coordinate system command
        AddItem("{5f914bd7-f9f1-4753-a603-d2c52a71b0c7}", 1) 'Add shp to TOC and workfolder command
        AddItem("{c6d4a1b8-ef1b-49eb-9c2d-691389528e66}", 1) 'Remove shp from TOC and workfolder command
        AddItem("{31d8e4a5-aa44-4bb7-b137-6a81aafc1dcb}", 1) 'Import from AutoCAD command
        AddItem("{a08d93a6-4b33-4ec0-8184-437bb0fd04e3}", 1) 'Split SHP command
        AddItem("{ab2748cd-ffa1-4d68-aa01-09fd165b3883}", 1) 'Safete copy of shp command

        'AddItem(New Guid("FBF8C3FB-0480-11D2-8D21-080009EE4E51"), 2) 'redo command
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "Input Data"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "InputDataMenu"
        End Get
    End Property
End Class


