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

<ComClass(cmd_6_6_VertexAtIntersection.ClassId, cmd_6_6_VertexAtIntersection.InterfaceId, cmd_6_6_VertexAtIntersection.EventsId), _
 ProgId("ArcArAz.cmd_6_6_VertexAtIntersection")> _
Public NotInheritable Class cmd_6_6_VertexAtIntersection
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "c2708dd4-4eac-4c0c-a9d3-a72282776128"
    Public Const InterfaceId As String = "43e71ddd-89e7-45bd-b30e-8714daf893be"
    Public Const EventsId As String = "a3f3bbb4-f642-4dd9-9736-7e65e51ca2d3"
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
        MyBase.m_caption = "Vertex at Entities Intersection"   'localizable text 
        MyBase.m_message = "Create a vertex at the intersection of selected features"   'localizable text 
        MyBase.m_toolTip = "Create a vertex at the intersection of selected features" 'localizable text 
        MyBase.m_name = "ArcArAz-Vertex_VertexAtIntersection"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_6_6_VertexAtIntersection.OnClick implementation
        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pSel As ISelection = pMap.FeatureSelection
        Dim pMapSel As IEnumFeature = pSel
        pMapSel.Reset()

        Dim i As Integer = 0
        Dim pfeat As IFeature = pMapSel.Next()
        Do Until pfeat Is Nothing
            i = i + 1
            pfeat = pMapSel.Next()
        Loop
        If i <> 2 Then
            MsgBox(i & " selected features." & vbCrLf & "Please, select only two features to intersect.")
            Exit Sub
        ElseIf i = 2 Then

            pMapSel.Reset()
            Dim pFeat1 As IFeature = pMapSel.Next()
            Dim pFeat2 As IFeature = pMapSel.Next()

            Dim pPoly1 As IGeometry
            Dim pGeomLine As IGeometryCollection = Nothing
            Dim pspFilter As ISpatialFilter = New SpatialFilter
            Dim pPointColl As IPointCollection
            Dim pPolycurve1 As IPolycurve
            Dim pPolycurve2 As IPolycurve
            Dim bSplitted As Boolean

            pspFilter.Geometry = pFeat1.ShapeCopy
            pspFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
            pPoly1 = pFeat1.ShapeCopy

            pPointColl = New Multipoint
            pPointColl = pFeat1.Shape
            pPolycurve1 = pPointColl


            Dim pPointResult As IPointCollection
            Dim pPoly2 As IGeometry

            pPoly2 = pFeat2.Shape

            pPolycurve2 = pPoly2

            Dim pGeomPolygon As IGeometryCollection
            pGeomPolygon = pPolycurve2

            For j = 0 To pGeomPolygon.GeometryCount - 1
                Dim pSegColl As ISegmentCollection
                pSegColl = New Path
                pSegColl.AddSegmentCollection(pGeomPolygon.Geometry(j))

                Dim pPolyline As IPolyline
                pPolyline = New Polyline

                pGeomLine = pPolyline
                pGeomLine.AddGeometry(pSegColl)
            Next
            pPointResult = GetIntersection(pPoly1, pGeomLine)
            For j = 0 To pPointResult.PointCount - 1
                Dim pInterPoint As IPoint
                pInterPoint = pPointResult.Point(j)
                pPolycurve1.SplitAtPoint(pInterPoint, True, False, bSplitted, 0, 0)
                pPolycurve2.SplitAtPoint(pInterPoint, True, False, bSplitted, 0, 0)
            Next
            pFeat1.Shape = pPolycurve1
            pFeat2.Shape = pPolycurve2
            If bSplitted = True Then
                pFeat1.Store()
                pFeat2.Store()
            End If
        End If
        MsgBox("Succeeded at " & Now, MsgBoxStyle.Information)
    End Sub
    Private Function GetIntersection(ByVal pIntersect As IGeometry, ByVal pOther As IPolyline) As IGeometry

        Dim pClone As IClone
        pClone = pIntersect.SpatialReference
        If Not pClone.IsEqual(pOther.SpatialReference) Then
            pOther.Project(pIntersect.SpatialReference)
        End If

        Dim pTopoOp As ITopologicalOperator
        pTopoOp = pOther
        pTopoOp.Simplify()

        pTopoOp = pIntersect

        Dim pGeomResult As IGeometry
        pGeomResult = pTopoOp.Intersect(pOther, esriGeometryDimension.esriGeometry0Dimension)

        If TypeOf pGeomResult Is IPointCollection Then
            Dim pPointColl As IPointCollection
            pPointColl = pGeomResult
        End If

        GetIntersection = pGeomResult
    End Function
End Class



