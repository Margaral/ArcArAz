Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem

<ComClass(cmd_6_7_VertexMultiIntersection.ClassId, cmd_6_7_VertexMultiIntersection.InterfaceId, cmd_6_7_VertexMultiIntersection.EventsId), _
 ProgId("ArcArAz.cmd_6_7_VertexMultiIntersection")> _
Public NotInheritable Class cmd_6_7_VertexMultiIntersection
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "6f09893c-e814-45f4-8831-6e772771ebd3"
    Public Const InterfaceId As String = "13c91144-899e-4e20-a82c-10da833e28aa"
    Public Const EventsId As String = "2716d3ca-4de5-4ffa-90b7-8d0c7be88e7d"
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
        MyBase.m_caption = "Vertex at Shapefiles Intersections"   'localizable text 
        MyBase.m_message = ""   'localizable text 
        MyBase.m_toolTip = "" 'localizable text 
        MyBase.m_name = "ArcArAz-Vertex_VertexMultiIntersection"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_6_7_VertexMultiIntersection.OnClick implementation

        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pLayerSet1 As ISet = pMxDoc.CurrentContentsView.SelectedItem
        Dim pLayerSet2 As ISet = pMxDoc.CurrentContentsView.SelectedItem
        'If pLayerSet1.Count <> 2 Then
        '    MsgBox("Please, select only 2 layers on TOC to intersect")
        '    Exit Sub
        'End If

        pLayerSet1.Reset()
        pLayerSet2.Reset()

        Dim pFLayer1 As IFeatureLayer = pLayerSet1.Next
        Dim pFLayer2 As IFeatureLayer = pLayerSet2.Next
        pFLayer2 = pLayerSet2.Next
        While Not pFLayer1 Is Nothing

            Dim psbar As IStatusBar = m_application.StatusBar
            Dim pPro As IStepProgressor = psbar.ProgressBar

            While Not pFLayer2 Is Nothing

                Dim pspFilter As ISpatialFilter = New SpatialFilter

                Dim pFCursor1 As IFeatureCursor = pFLayer1.FeatureClass.Search(Nothing, False)

                Dim pFeat1 As IFeature = pFCursor1.NextFeature

                Dim pPoly1 As IGeometry
                Dim pGeomLine As IGeometryCollection = Nothing

                Dim pPointColl As IPointCollection
                Dim pPolycurve1 As IPolycurve
                Dim pPolycurve2 As IPolycurve
                'Dim bProject As Boolean, bCreatePart As Boolean, 
                Dim bSplitted As Boolean
                'Dim lNewPart As Long, lNewSeg As Long
                Dim pClone1 As IClone
                Dim pFCursor2 As IFeatureCursor
                Dim pFeat2 As IFeature


                pPro.MinRange = 1
                pPro.MaxRange = pFLayer1.FeatureClass.FeatureCount(Nothing)
                pPro.StepValue = pFLayer1.FeatureClass.FeatureCount(Nothing) / 100
                pPro.Message = pFLayer1.Name
                pPro.Step()
                pPro.Show()

                While Not pFeat1 Is Nothing
                    pspFilter.Geometry = pFeat1.ShapeCopy
                    pspFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects
                    pPoly1 = pFeat1.ShapeCopy

                    pPointColl = New Multipoint
                    pPointColl = pFeat1.ShapeCopy
                    pClone1 = pPointColl
                    pPolycurve1 = pClone1.Clone
                    pFCursor2 = pFLayer2.FeatureClass.Search(pspFilter, False)
                    pFeat2 = pFCursor2.NextFeature
                    Dim pPointResult As IPointCollection
                    Dim pPoly2 As IGeometry
                    Do Until pFeat2 Is Nothing
                        pPoly2 = pFeat2.ShapeCopy

                        Dim pClone2 As IClone = pPoly2
                        pPolycurve2 = pClone2.Clone

                        Dim pGeomPolygon As IGeometryCollection = pClone2.Clone

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
                        pFeat2 = pFCursor2.NextFeature
                    Loop
                    pFeat1 = pFCursor1.NextFeature
                    pPro.Step()

                End While
                pFLayer2 = pLayerSet2.Next
            End While
            pPro.Hide()
            pLayerSet2.Reset()
            pFLayer1 = pLayerSet1.Next
        End While

        MsgBox("Succeeded at " & Now, MsgBoxStyle.Information)
    End Sub



    Private Function GetIntersection(ByVal pIntersect As IGeometry, ByVal pOther As IPolyline) As IGeometry

        Dim pClone As IClone = pIntersect.SpatialReference
        If Not pClone.IsEqual(pOther.SpatialReference) Then
            pOther.Project(pIntersect.SpatialReference)
        End If

        Dim pTopoOp As ITopologicalOperator = pOther
        pTopoOp.Simplify()

        pTopoOp = pIntersect

        Dim pGeomResult As IGeometry = pTopoOp.intersect(pOther, esriGeometryDimension.esriGeometry0Dimension)

        If TypeOf pGeomResult Is IPointCollection Then
            Dim pPointColl As IPointCollection
            pPointColl = pGeomResult
        End If

        GetIntersection = pGeomResult
    End Function

End Class



