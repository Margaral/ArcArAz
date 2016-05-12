Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(menu2EntityProperties.ClassId, menu2EntityProperties.InterfaceId, menu2EntityProperties.EventsId), _
 ProgId("ArcArAz.EntityProperties")> _
Public NotInheritable Class menu2EntityProperties
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
    Public Const ClassId As String = "1d5323f8-99c0-41d9-b3e7-c9788fe0cc88"
    Public Const InterfaceId As String = "1f0a1932-c5fa-4170-bd7a-96471cd63c30"
    Public Const EventsId As String = "63af3fa4-f416-4d66-9b64-4b65c5e843b3"
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
        AddItem("{8bac7d2f-79bd-4453-be92-1a1af544e9b3}", 1) 'Remove interior holes command
        AddItem("{4d21f6b7-188e-42c3-a189-5251775c7a82}", 1) 'Line to Polygon
        AddItem("{4a15c147-1c8b-455b-9fd6-267cf824c138}", 1) 'Multipart to singlepart command
        AddItem("{8bb287ed-872e-40fd-a5e3-242201ff5e3c}", 1) 'Start Edit Session command
        AddItem("{88662805-e92c-44de-85bf-96ea683748fe}", 1) 'Rename Entities
        AddItem("{d10aaab6-dac7-4945-a74d-f339280a9a29}", 1) 'Copy feature

        'AddItem(New Guid("FBF8C3FB-0480-11D2-8D21-080009EE4E51"), 2) 'redo command
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "Entity Properties"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "EntityPropertiesMenu"
        End Get
    End Property
End Class


