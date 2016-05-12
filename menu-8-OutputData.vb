Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports System.Runtime.InteropServices

<ComClass(menu8OutputData.ClassId, menu8OutputData.InterfaceId, menu8OutputData.EventsId), _
 ProgId("ArcArAz.menuOutputData")> _
Public NotInheritable Class menu8OutputData
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
    Public Const ClassId As String = "e3f068fd-bdf7-4a77-9c6b-845998d51c2c"
    Public Const InterfaceId As String = "baae17c1-b2e2-45a8-a630-8e099b2a7e79"
    Public Const EventsId As String = "1c6d8102-4b66-4144-8da2-0696fdb499a3"
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
        'AddItem(New Guid("FBF8C3FB-0480-11D2-8D21-080009EE4E51"), 2) 'redo command

        AddItem("{b95e349b-81cd-4701-b320-88eab41a9d7f}", 1) '8_4_Shp with mesh nodes command
        AddItem("{f8776c4c-ef49-45bf-a9ad-31f77c0c1eab}", 1) '8_3_Shp with mesh elements command

        AddItem("{ac2346ae-ab20-44f7-babc-3610f15e6d45}", 1) 'Raster / TIN Piezo command
        AddItem("{25ea907a-b137-4302-8afa-1a63a7e727ea}", 1) 'Raster / TIN with descensos command

        AddItem("{bad0c77e-5cc1-4d18-ad68-1b3614a99d62}", 1) 'PLT Viewer .exe 
        AddItem("{83c602e8-28a4-4683-a35e-7ba99bd00466}", 1) 'DT2VDTK
        AddItem("{642e8a59-538f-46ee-9d15-2a6ff9c0d9d0}", 1) 'parametros
        AddItem("{bae94db5-6b52-4824-9e17-3c2b7db668c7}", 1) 'Particle track tool
    End Sub

    Public Overrides ReadOnly Property Caption() As String
        Get
            'TODO: Replace bar caption
            Return "Output Data"
        End Get
    End Property

    Public Overrides ReadOnly Property Name() As String
        Get
            'TODO: Replace bar ID
            Return "OutputDataMenu"
        End Get
    End Property
End Class


