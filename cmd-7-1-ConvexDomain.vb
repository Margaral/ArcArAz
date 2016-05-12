Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.esriSystem

<ComClass(cmd_7_1_ConvexDomain.ClassId, cmd_7_1_ConvexDomain.InterfaceId, cmd_7_1_ConvexDomain.EventsId), _
 ProgId("ArcArAz.cmd_7_1_ConvexDomain")> _
Public NotInheritable Class cmd_7_1_ConvexDomain
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ac0ba63b-614c-4645-a86a-6cf0227c16e0"
    Public Const InterfaceId As String = "399af1e9-efce-459b-8237-26f6f0944ce1"
    Public Const EventsId As String = "77087d69-58dc-47b6-ad1b-eeaf238a6323"
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
        MyBase.m_category = "ArcArAz-Domain"  'localizable text 
        MyBase.m_caption = "Convex Domain"   'localizable text 
        MyBase.m_message = "Create a convex domain of selected features"   'localizable text 
        MyBase.m_toolTip = "Create a convex domain of selected features" 'localizable text 
        MyBase.m_name = "ArcArAz-Domain_ConvexDomain"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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

        Dim pMxDoc As IMxDocument = m_application.Document

        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pGeomColl As IGeometryCollection = New Polyline
        Dim pInGeomColl As IGeometryCollection
        Dim i As Integer
        For i = 0 To pMap.LayerCount - 1
            On Error Resume Next
            Dim pFeatSel As IFeatureSelection
            pFeatSel = pMap.Layer(i)

            Dim pSelSet As ISelectionSet
            pSelSet = pFeatSel.SelectionSet

            Dim pInFCursor As IFeatureCursor = Nothing

            If pSelSet.Count <> 0 Then
                pFeatSel.SelectionSet.Search(Nothing, True, pInFCursor)
                Dim pFeat As IFeature
                pFeat = pInFCursor.NextFeature

                Dim j As Integer
                For j = 0 To pSelSet.Count - 1
                    If pFeat.Shape.GeometryType = 3 Then

                        pInGeomColl = pFeat.Shape
                        pGeomColl.AddGeometry(pInGeomColl.Geometry(0))
                        pFeat = pInFCursor.NextFeature

                    ElseIf pFeat.Shape.GeometryType = 4 Then

                        Dim pPoly As IPolygon
                        pPoly = pFeat.Shape

                        Dim pClone As IClone
                        pClone = pPoly

                        Dim pGeomPolygon As IGeometryCollection
                        pGeomPolygon = pClone.Clone

                        For k = 0 To pGeomPolygon.GeometryCount - 1
                            Dim pSegColl As ISegmentCollection
                            pSegColl = New Path
                            pSegColl.AddSegmentCollection(pGeomPolygon.Geometry(k))

                            pGeomColl.AddGeometry(pSegColl)
                        Next
                        pFeat = pInFCursor.NextFeature
                    End If
                Next
            End If
        Next

        Dim pInLayer As IFeatureLayer
        pInLayer = pMap.Layer(0)

        Dim pInFeatClass As IFeatureClass
        pInFeatClass = pInLayer.FeatureClass

        Dim pDataset As IDataset
        pDataset = pInFeatClass

        Dim strFolder As String
        strFolder = pDataset.Workspace.PathName

        Dim strName As String
        strName = pDataset.Name & "_Domain"

        Dim pWSF As IWorkspaceFactory
        pWSF = New ShapefileWorkspaceFactory

        Dim pFeatWS As IFeatureWorkspace
        Dim pFeatDS As IFeatureClass

        If pWSF.IsWorkspace(strFolder) Then
            pFeatWS = pWSF.OpenFromFile(strFolder, 0)
            On Error Resume Next
            pFeatDS = pFeatWS.OpenFeatureClass(strName)
            Dim pFeatDataSet As IDataset
            pFeatDataSet = pFeatDS
            pFeatDataSet.Delete()
        End If

        Const strShapeFieldName As String = "Shape"

        ' Open the folder to contain the shapefile as a workspace
        Dim pFWS As IFeatureWorkspace
        Dim pWorkspaceFactory As IWorkspaceFactory
        pWorkspaceFactory = New ShapefileWorkspaceFactory
        pFWS = pWorkspaceFactory.OpenFromFile(strFolder, 0)

        ' Set up a simple fields collection
        Dim pFields As IFields
        Dim pFieldsEdit As IFieldsEdit
        pFields = New Fields
        pFieldsEdit = pFields

        Dim pField As IField
        Dim pFieldEdit As IFieldEdit

        ' Make the shape field
        ' it will need a geometry definition, with a spatial reference
        pField = New Field
        pFieldEdit = pField
        pFieldEdit.Name_2 = strShapeFieldName
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry

        Dim pGeomDef As IGeometryDef
        Dim pGeomDefEdit As IGeometryDefEdit
        pGeomDef = New GeometryDef
        pGeomDefEdit = pGeomDef
        With pGeomDefEdit
            .GeometryType_2 = esriGeometryType.esriGeometryPolygon
            .SpatialReference_2 = New UnknownCoordinateSystem 'pInFeatClass.GetFeature(0).Shape.SpatialReference
        End With
        pFieldEdit.GeometryDef_2 = pGeomDef
        pFieldsEdit.AddField(pField)

        ' Add another miscellaneous text field
        pField = New Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 30
            .Name_2 = "Text"
            .Type_2 = esriFieldType.esriFieldTypeString
        End With
        pFieldsEdit.AddField(pField)

        ' Create the shapefile
        ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
        Dim pOutFeatClass As IFeatureClass
        pOutFeatClass = pFWS.CreateFeatureClass(strName, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, strShapeFieldName, "")

        Dim pTopoOp As ITopologicalOperator
        pTopoOp = pGeomColl
        Dim pPolygon As IPolygon4
        pTopoOp.Simplify()
        pPolygon = pTopoOp.ConvexHull

        Dim pOutFeat As IFeature
        pOutFeat = pOutFeatClass.CreateFeature
        pOutFeat.Shape = pPolygon
        pOutFeat.Store()

        For s = 1 To pMap.LayerCount - 1
            If pMap.Layer(s).Name = pOutFeatClass.AliasName Then
                pMap.DeleteLayer(pMap.Layer(s))
            End If
        Next

        Dim pOutFLayer As IFeatureLayer
        pOutFLayer = New FeatureLayer
        pOutFLayer.FeatureClass = pOutFeatClass
        pOutFLayer.Name = pOutFLayer.FeatureClass.AliasName

        pMap.AddLayer(pOutFLayer)
        pMap.MoveLayer(pOutFLayer, (pMxDoc.FocusMap.LayerCount))

        pMxDoc.ActiveView.ContentsChanged()
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()



    End Sub

End Class



