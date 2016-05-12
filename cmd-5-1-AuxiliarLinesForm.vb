Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.DataSourcesFile
Imports System.IO
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem

Public Class cmd_5_1_AuxiliarLinesForm

    Private m_application As IApplication

    Public Property Application() As IApplication
        Get
            Return m_application
        End Get
        Set(ByVal value As IApplication)
            m_application = value
        End Set
    End Property

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim nanillos As Integer = CUInt(TextBox1.Text.ToString)
        Dim distIni As Double = CDbl(TextBox2.Text.ToString)
        Dim distFin As Double = CDbl(TextBox3.Text.ToString)
        Dim nlados As Integer = CUInt(TextBox4.Text.ToString)
        Dim ExpdistIni As Double = Math.Log(distIni)
        Dim ExpdistFin As Double = Math.Log(distFin)
        Dim Intervalo As Double = (ExpdistFin - ExpdistIni) / nanillos

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pSel As ISelection = pMap.FeatureSelection
        Dim pMapSel As IEnumFeature = pSel
        pMapSel.Reset()

        'Get the path of the .mxd document
        Dim templates As ESRI.ArcGIS.Framework.ITemplates = m_application.Templates
        Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra
        'If System.IO.Directory.Exists(folderName & "\Auxiliar Lines for Wells\") = True Then
        '    For Each File As String In System.IO.Directory.GetFiles(folderName & "\Auxiliar Lines for Wells\")
        '        System.IO.File.Delete(File)
        '    Next
        '    System.IO.Directory.Delete(folderName & "\Auxiliar Lines for Wells\")
        'End If
        System.IO.Directory.CreateDirectory(folderName & "\Auxiliar Lines for Wells\")
        folderName = folderName & "\Auxiliar Lines for Wells\"
        Dim workspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactoryClass
        Dim featureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = CType(workspaceFactory.OpenFromFile(folderName, 0), ESRI.ArcGIS.Geodatabase.IFeatureWorkspace) ' Explicit Cast

        CreateFeatureClassAndLayer(folderName, "AuxiliarLines")
        Dim strPathFC As String = folderName & "AuxiliarLines.shp"
        Dim pFeatureLayer As IFeatureLayer = pMxDoc.FocusMap.Layer(pMxDoc.FocusMap.LayerCount - 1)
        Dim anillosFC As IFeatureClass = pFeatureLayer.FeatureClass 'featureWorkspace.OpenFeatureClass(strPathFC)
        'Dim featureLayer As IFeatureLayer = New FeatureLayerClass
        'featureLayer.FeatureClass = anillosFC
        'featureLayer.Name = anillosFC.AliasName
        'featureLayer.Visible = True
        'pMxDoc.FocusMap.AddLayer(featureLayer)

        Dim pfeat As IFeature = pMapSel.Next()

        Do Until pfeat Is Nothing
            If pfeat.Shape.GeometryType = esriGeometryType.esriGeometryPoint Then
                Dim pCenterPoint As ESRI.ArcGIS.Geometry.IPoint = pfeat.ShapeCopy

                For i As Integer = 0 To nanillos
                    Dim radio As Double = Math.Exp(ExpdistIni + Intervalo * i)

                    If nlados = 0 Then
                        Dim preRadio As Double = Math.Exp(ExpdistIni + Intervalo * (i - 1))
                        Dim longLado As Double = radio - preRadio
                        nlados = 2 * Math.PI * radio / longLado
                    End If

                    Dim angulo As Double = 2 * Math.PI / nlados
                    Dim pPointColl As IPointCollection = New Multipoint

                    If i = 0 Then
                        For k As Integer = 1 To nlados
                            Dim pPoint2 As ESRI.ArcGIS.Geometry.IPoint = New Point
                            pPoint2.X = pCenterPoint.X + radio * Math.Cos(angulo * k)
                            pPoint2.Y = pCenterPoint.Y + radio * Math.Sin(angulo * k)

                            Dim pLine As ILine = New Line
                            pLine.PutCoords(pCenterPoint, pPoint2)

                            Dim pPath As ISegmentCollection = New ESRI.ArcGIS.Geometry.Path
                            pPath.AddSegment(pLine)

                            Dim pPolyline As IPolyline = New Polyline
                            Dim pGeomColl As IGeometryCollection = pPolyline
                            pGeomColl.AddGeometry(pPath)

                            Dim pFeatLine As IFeature = anillosFC.CreateFeature
                            pFeatLine.Shape = pGeomColl
                            pFeatLine.Value(3) = "Ring_" & pfeat.OID.ToString & "_0_" & k
                            pFeatLine.Store()
                        Next
                    End If

                    For j As Integer = 1 To nlados
                        Dim pPoint As ESRI.ArcGIS.Geometry.IPoint = New Point
                        pPoint.X = pCenterPoint.X + radio * Math.Cos(angulo * j)
                        pPoint.Y = pCenterPoint.Y + radio * Math.Sin(angulo * j)
                        pPointColl.AddPoint(pPoint)
                    Next

                    Dim pTopoOp As ITopologicalOperator = pPointColl
                    pTopoOp.Simplify()
                    Dim pGeom As IGeometry = pTopoOp.ConvexHull

                    Dim pPolygon As IPolygon = pGeom
                    Dim pRing As ISegmentCollection = pPolygon
                    Dim pPathPolyline As ISegmentCollection = New Polyline
                    pPathPolyline.AddSegmentCollection(pRing)

                    Dim pPolylineHull As IPolyline = pPathPolyline

                    Dim pFeatRing As IFeature = anillosFC.CreateFeature
                    pFeatRing.Shape = pPolylineHull
                    pFeatRing.Value(2) = radio
                    pFeatRing.Value(3) = "Ring_" & pfeat.OID.ToString & "_" & i
                    pFeatRing.Store()
                Next
            ElseIf pfeat.Shape.GeometryType = esriGeometryType.esriGeometryPolyline Or pfeat.Shape.GeometryType = esriGeometryType.esriGeometryPolygon Then
                For i As Integer = 0 To nanillos
                    Dim radio As Double = Math.Exp(ExpdistIni + Intervalo * i)

                    Dim topologicalOperator As ITopologicalOperator = pfeat.Shape
                    Dim pPolygon As IPolygon = New Polygon
                    pPolygon = topologicalOperator.Buffer(radio)
                    Dim pPolyline As IPolyline = PolygonToPolyline(pPolygon)

                    If nlados = 0 Then
                        Dim preRadio As Double = Math.Exp(ExpdistIni + Intervalo * (i - 1))
                        Dim longLado As Double = radio - preRadio
                        'nlados = 2 * Math.PI * radio / longLado
                        pPolyline.Densify(longLado, 0)
                    End If

                    Dim pFeatRing As IFeature = anillosFC.CreateFeature
                    pFeatRing.Shape = pPolyline
                    pFeatRing.Value(2) = radio
                    pFeatRing.Value(3) = "Ring_" & pfeat.OID.ToString & "_" & i
                    pFeatRing.Store()
                Next
            End If
            pfeat = pMapSel.Next()
        Loop
        Me.Close()
        pMxDoc.ActiveView.Refresh()
    End Sub
    Private Function PolygonToPolyline(ByRef pPolygon As IPolygon) As IGeometryCollection
        PolygonToPolyline = New Polyline
        Dim pGeoms_Polygon As IGeometryCollection, pClone As IClone
        pClone = pPolygon
        pGeoms_Polygon = pClone.Clone
        Dim i As Long, pSegs_Path As ISegmentCollection
        For i = 0 To pGeoms_Polygon.GeometryCount - 1
            pSegs_Path = New ESRI.ArcGIS.Geometry.Path
            pSegs_Path.AddSegmentCollection(pGeoms_Polygon.Geometry(i))
            PolygonToPolyline.AddGeometry(pSegs_Path)
        Next i
    End Function

    Public Sub CreateFeatureClassAndLayer(ByVal strPathGeneral As String, ByVal strName As String)

        strPathGeneral = folderName

        Const strShapeFieldName As String = "Shape"

        ' Open the folder to contain the shapefile as a workspace

        Dim pWorkspaceFactory As IWorkspaceFactory = New ShapefileWorkspaceFactory
        Dim pFWS As IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(strPathGeneral, 0)

        ' Delete an old feature class
        If System.IO.File.Exists(strPathGeneral & "\" & strName & ".shp") Then
            Dim pDelFC As IFeatureClass = pFWS.OpenFeatureClass(strName)
            Dim pDataset As IDataset = pDelFC
            pDataset.Delete()
        End If

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
            .GeometryType_2 = esriGeometryType.esriGeometryPolyline
            .HasZ_2 = False
            .SpatialReference_2 = New UnknownCoordinateSystem 'pInFeatClass.GetFeature(0).Shape.SpatialReference
        End With
        pFieldEdit.GeometryDef_2 = pGeomDef
        pFieldsEdit.AddField(pField)

        ' Add another miscellaneous text field
        pField = New Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 10
            .Name_2 = "Distance"
            .Type_2 = esriFieldType.esriFieldTypeDouble
        End With
        pFieldsEdit.AddField(pField)

        ' Add another miscellaneous text field
        pField = New Field
        pFieldEdit = pField
        With pFieldEdit
            .Length_2 = 20
            .Name_2 = "Name"
            .Type_2 = esriFieldType.esriFieldTypeString
        End With
        pFieldsEdit.AddField(pField)


        ' Create the shapefile
        ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
        Dim pFeatClass As IFeatureClass = pFWS.CreateFeatureClass(strName, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, strShapeFieldName, "")

        'Add shapefile to TOC
        'Create a new ShapefileWorkspaceFactory object and open a shapefile folder

        Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(strPathGeneral, 0)

        'Create a new FeatureLayer and assign a shapefile to it
        Dim pFeatureLayer As IFeatureLayer = New FeatureLayer
        pFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(strName)
        pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName

        'Add the FeatureLayer to the focus map
        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        pMap.AddLayer(pFeatureLayer)
        pMap.MoveLayer(pFeatureLayer, (pMxDoc.FocusMap.LayerCount))

    End Sub


End Class