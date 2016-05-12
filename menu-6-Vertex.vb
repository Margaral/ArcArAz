Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(menu6Vertex.ClassId, menu6Vertex.InterfaceId, menu6Vertex.EventsId), _
 ProgId("ArcArAz.menuVertex")> _
Public NotInheritable Class menu6Vertex
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
    Public Const ClassId As String = "d1fa34bf-1fbe-4b98-be81-b165a4b7e21b"
    Public Const InterfaceId As String = "6c7f84f0-0cb8-4ac7-bfaf-72f5e9ddf963"
    Public Const EventsId As String = "c064d74f-c00a-404f-96f4-ac87f7958e40"
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
        AddItem("{054ca18d-d0a0-421b-8aeb-38a5b716fcfd}", 1) 'Select entity to submit command
        AddItem("{7b00855c-7232-415c-bcd2-650863a058ff}", 1) 'Visualize vertex entities command
        AddItem("{593fb203-f05f-4bc2-a474-6d925d7e8233}", 1) 'Vertex count command
        AddItem("{8c7a3b77-0e21-4009-b9b6-591350f0b4ff}", 1) 'Reduce vertex number command
        AddItem("{c2708dd4-4eac-4c0c-a9d3-a72282776128}", 1) 'Vertex at intersection command
        AddItem("{6f09893c-e814-45f4-8831-6e772771ebd3}", 1) 'Vertex at multiintersection command
        AddItem("{91df2c2c-2111-4cb3-b5d1-514426503bcf}", 1) 'Increase vertex number submenu


        'AddItem("{FBF8C3FB-0480-11D2-8D21-080009EE4E51}", 1) 'undo command
        'AddItem(New Guid("FBF8C3FB-0480-11D2-8D21-080009EE4E51"), 2) 'redo command
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "Vertex"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "VertexMenu"
        End Get
    End Property
End Class


