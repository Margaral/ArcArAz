Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(toolbarArcArAzTools.ClassId, toolbarArcArAzTools.InterfaceId, toolbarArcArAzTools.EventsId), _
 ProgId("ArcArAz.ArcArAzTools")> _
Public NotInheritable Class toolbarArcArAzTools
    Inherits BaseToolbar

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
    Public Const ClassId As String = "7a456fac-51c8-4e6c-b736-df8272df447a"
    Public Const InterfaceId As String = "e0d25484-cc4d-4b70-b554-2c4c24be1ca6"
    Public Const EventsId As String = "132e6987-eb33-49e8-93a8-af9a5a467007"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()

        '
        'TODO: Define your toolbar here by adding items
        '
        'AddItem("esriArcMapUI.ZoomOutTool")
        'BeginGroup() 'Separator
        AddItem("{3c4884c3-64ff-4946-9284-c75cd1044df0}", 1) 'InputDataMenu
        BeginGroup() 'Separator
        AddItem("{1d5323f8-99c0-41d9-b3e7-c9788fe0cc88}", 1) 'EntityPropertiesMenu
        BeginGroup() 'Separator
        AddItem("{c996a507-89eb-4db2-bf8e-579498695488}", 1) 'SpatialFieldsMenu
        BeginGroup() 'Separator
        AddItem("{c1c26f6f-5638-4e69-8efa-7ed487185355}", 1) 'TemporalSeriesMenu
        BeginGroup()
        AddItem("{bc775037-21ec-4414-9d57-82eb8479f4d1}", 1) 'RiversMenu
        BeginGroup() 'Separator
        AddItem("{a378b01e-bf01-4a28-aadf-28e7eb78ed5a}", 1) 'WellsMenu
        BeginGroup() 'Separator
        AddItem("{29c6d54b-d518-472a-8869-32a7e99cf205}", 1) 'DomainMenu
        BeginGroup() 'Separator
        AddItem("{d1fa34bf-1fbe-4b98-be81-b165a4b7e21b}", 1) 'VertexMenu
        BeginGroup() 'Separator
        AddItem("{76e23a88-8e1c-4889-9bfc-d001b03fb559}", 1) 'To VT Menu
        BeginGroup()
        AddItem("{e3f068fd-bdf7-4a77-9c6b-845998d51c2c}", 1) 'OutputMenu
        BeginGroup()
        AddItem("{4ba85aef-99d8-4c6b-a742-cd45a53a0acf}", 1) 'Help



        'AddItem(New Guid("FBF8C3FB-0480-11D2-8D21-080009EE4E51"), 2) 'redo command
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "ArcArAz"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "ArcArAzToolsToolBar"
        End Get
    End Property
End Class
