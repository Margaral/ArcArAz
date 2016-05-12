Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem

<ComClass(cmd_6_1_VisualizeVertexEntities.ClassId, cmd_6_1_VisualizeVertexEntities.InterfaceId, cmd_6_1_VisualizeVertexEntities.EventsId), _
 ProgId("ArcArAz.cmd_6_1_VisualizeVertexEntities")> _
Public NotInheritable Class cmd_6_1_VisualizeVertexEntities
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "7b00855c-7232-415c-bcd2-650863a058ff"
    Public Const InterfaceId As String = "8f752583-280b-4271-b9ad-34ebe3188c71"
    Public Const EventsId As String = "1660c42d-1611-4860-bfe6-e5139dfbabf4"
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
        MyBase.m_category = "ArcArAz-Vertex"  'localizable text 
        MyBase.m_caption = "Visualize Vertex Entities"   'localizable text 
        MyBase.m_message = "Create graphics representing vertex of selected entity"   'localizable text 
        MyBase.m_toolTip = "Create graphics representing vertex of selected entity" 'localizable text 
        MyBase.m_name = "ArcArAz-Vertex_VisualizeVertexEntities"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_6_1_VisualizeVertexEntities.OnClick implementation
        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pPointColl As IPointCollection
        Dim pGraphics As IGraphicsContainer = pMxDoc.FocusMap
        Dim pActiveView As IActiveView = pGraphics

        Dim i As Integer
        For i = 0 To pMap.LayerCount - 1

            Dim pFeatSel As IFeatureSelection = pMap.Layer(i)
            Dim pSelSet As ISelectionSet = pFeatSel.SelectionSet
            Dim pInFCursor As IFeatureCursor = Nothing

            If pSelSet.Count <> 0 Then
                pFeatSel.SelectionSet.Search(Nothing, True, pInFCursor)
                Dim pFeat As IFeature
                pFeat = pInFCursor.NextFeature

                Dim j As Integer
                For j = 0 To pSelSet.Count - 1
                    pPointColl = pFeat.Shape
                    For k = 0 To pPointColl.PointCount - 1
                        Dim pPoint As IPoint = pPointColl.Point(k)
                        Dim pElement As IElement = New MarkerElement
                        pElement.Geometry = pPoint
                        pGraphics.AddElement(pElement, 0)
                        pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, pElement, Nothing)
                    Next
                    pFeat = pInFCursor.NextFeature

                Next
            End If
        Next

    End Sub
End Class



