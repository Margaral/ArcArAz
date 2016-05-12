Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(menu_10_ToVT.ClassId, menu_10_ToVT.InterfaceId, menu_10_ToVT.EventsId), _
 ProgId("ArcArAz.menu_10_ToVT")> _
Public NotInheritable Class menu_10_ToVT
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
    Public Const ClassId As String = "76e23a88-8e1c-4889-9bfc-d001b03fb559"
    Public Const InterfaceId As String = "204cec9a-1000-4c12-806e-2f98fe37c9cb"
    Public Const EventsId As String = "c8618d79-5dcd-448a-9df9-608d16af696e"
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
        'AddItem("{FBF8C3FB-0480-11D2-8D21-080009EE4E51}", 1) 'undo command
        'AddItem(New Guid("FBF8C3FB-0480-11D2-8D21-080009EE4E51"), 2) 'redo command

        AddItem("{91b5ce78-cbd4-4055-a97d-99dc5bc7423f}", 1) 'Sharp edges entity command
        AddItem("{62ffcef0-76b8-4dae-9f26-55c4e9d0a62d}", 1) 'Sharp edges command
        AddItem("{658e39fc-1864-43b1-ac8d-3b5163e795ad}", 1) 'Measure angle tool
        BeginGroup()
        AddItem("{44db8f26-d34e-493d-aa13-0300a7172a3e}", 1) 'Topology modelbuider
        BeginGroup()
        AddItem("{3728685e-5b64-4309-af19-de18a32e478b}", 1) 'Central Coordinates command

    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "Meshing"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "Meshing"
        End Get
    End Property
End Class


