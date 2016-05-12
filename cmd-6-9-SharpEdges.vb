Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

<ComClass(cmd_6_9_SharpEdges.ClassId, cmd_6_9_SharpEdges.InterfaceId, cmd_6_9_SharpEdges.EventsId), _
 ProgId("ArcArAz.cmd_6_9_SharpEdges")> _
Public NotInheritable Class cmd_6_9_SharpEdges
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "62ffcef0-76b8-4dae-9f26-55c4e9d0a62d"
    Public Const InterfaceId As String = "363597a4-09bc-4cac-9d83-88f47f0adc26"
    Public Const EventsId As String = "9a51bb7c-7312-4ca8-82bf-b606bad99871"
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
        MyBase.m_caption = "Eliminate Sharp Edges"   'localizable text 
        MyBase.m_message = "Eliminate sharp edges between entities"   'localizable text 
        MyBase.m_toolTip = "Eliminate sharp edges between entities" 'localizable text 
        MyBase.m_name = "ArcArAz-Vertex_SharpEdges"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_6_9_SharpEdges.OnClick implementation

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
            MsgBox(i & " selected features." & vbCrLf & "Please, select only two features to adjust.")
            Exit Sub
        ElseIf i = 2 Then

            pMapSel.Reset()
            Dim pFeat1 As IFeature = pMapSel.Next()
            Dim pFeat2 As IFeature = pMapSel.Next()

            Dim pPoly1 As IPolygon = pFeat1.Shape
            Dim pPoly2 As IPolygon = pFeat2.Shape

            Dim SeTocan As Boolean = GetPolygonIntersection(pPoly1, pPoly2)

            If SeTocan = True Then
                Dim pPointColl1 As IPointCollection = pPoly1
                Dim pPointColl2 As IPointCollection = pPoly2

                For j = 0 To pPointColl1.PointCount - 2
                    Dim pOutVertex1 As IPoint = pPointColl1.Point(j)

                    For k = 0 To pPointColl2.PointCount - 2
                        Dim pRelOp As IRelationalOperator = pOutVertex1
                        Dim SonIguales As Boolean = pRelOp.Equals(pPointColl2.Point(k))

                        If SonIguales = True Then
                            Dim pPointPrevious1 As IPoint
                            Dim pPointNext1 As IPoint
                            Dim pPointPrevious2 As IPoint
                            Dim pPointNext2 As IPoint

                            If j = 0 Then
                                pPointPrevious1 = pPointColl1.Point(pPointColl1.PointCount - 2)
                                pPointNext1 = pPointColl1.Point(1)
                            ElseIf j = pPointColl1.PointCount - 2 Then
                                pPointPrevious1 = pPointColl1.Point(j - 1)
                                pPointNext1 = pPointColl1.Point(0)
                            Else
                                pPointPrevious1 = pPointColl1.Point(j - 1)
                                pPointNext1 = pPointColl1.Point(j + 1)
                            End If

                            If k = 0 Then
                                pPointPrevious2 = pPointColl2.Point(pPointColl1.PointCount - 2)
                                pPointNext2 = pPointColl2.Point(1)
                            ElseIf k = pPointColl2.PointCount - 2 Then
                                pPointPrevious2 = pPointColl2.Point(k - 1)
                                pPointNext2 = pPointColl2.Point(0)
                            Else
                                pPointPrevious2 = pPointColl2.Point(k - 1)
                                pPointNext2 = pPointColl2.Point(k + 1)
                            End If

                            Dim grados1 As Double = GetAngulo(pOutVertex1, pPointNext1, pPointNext2)
                            Dim grados2 As Double = GetAngulo(pOutVertex1, pPointNext1, pPointPrevious2)
                            Dim grados3 As Double = GetAngulo(pOutVertex1, pPointPrevious1, pPointNext2)
                            Dim grados4 As Double = GetAngulo(pOutVertex1, pPointPrevious1, pPointPrevious2)

                            MsgBox(grados1 & " " & grados2 & " " & grados3 & " " & grados4)
                        End If
                    Next
                Next
            End If
        End If
        'MsgBox("Succeeded at " & Now, MsgBoxStyle.Information)
        pMxDoc.ActiveView.Refresh()
    End Sub

    Public Function GetAngulo(ByVal puntoC As IPoint, ByVal punto1 As IPoint, ByVal punto2 As IPoint) As Double
        Dim m1 As Double = (punto1.Y - puntoC.Y) / (punto1.X - puntoC.X)
        Dim m2 As Double = (punto2.Y - puntoC.Y) / (punto2.X - puntoC.X)
        Dim beta As Double = 1 / Math.Tan((m2 - m1) / (1 + m2 * m1))

        Dim moduloA As Double = Math.Sqrt((puntoC.X - punto1.X) ^ 2 + (puntoC.Y - punto1.Y) ^ 2)
        Dim moduloB As Double = Math.Sqrt((puntoC.X - punto2.X) ^ 2 + (puntoC.Y - punto2.Y) ^ 2)
        Dim prodEscalar As Double = (puntoC.X - punto1.X) * (puntoC.X - punto2.X) + (puntoC.Y - punto1.Y) * (puntoC.Y - punto2.Y)
        Dim alfa As Double = Math.Acos(prodEscalar / (moduloA * moduloB)) * 180 / Math.PI

        Return alfa
    End Function

    Public Function GetPolygonIntersection(ByVal pEntity1 As IPolygon, ByVal pEntity2 As IPolygon) As Boolean

        Dim pGeom As IGeometry
        Dim pTopo As ITopologicalOperator = pEntity1

        If Not pTopo.IsSimple Then pTopo.Simplify()
        pGeom = pTopo.Intersect(pEntity2, esriGeometryDimension.esriGeometry0Dimension)
        If pGeom.IsEmpty Then
            pGeom = pTopo.Intersect(pEntity2, esriGeometryDimension.esriGeometry1Dimension)
            If pGeom.IsEmpty Then
                pGeom = pTopo.Intersect(pEntity2, esriGeometryDimension.esriGeometry2Dimension)
                If pGeom.IsEmpty Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Else
            Return True
        End If

    End Function
End Class

'Sub IEnumVertexExample()
'    Dim pPointCollection As IPointCollection, pEnumVertex As IEnumVertex2
'    Dim pOutVertex As IPoint, lOutPart As Long, lOutVertex As Long
'    Dim pClonedEnumVertex As IEnumVertex2, pQueryVertex As IPoint
'    Dim OutWKSPoint As WKSPoint, pMAware As IMAware
'    Dim pZAware As IZAware, pPtIdAware As IPointIDAware, i As Long
'    'Get a polyline object created in CreateMultipartPolyline
'    'and QI for IPointCollection
'    pPointCollection = CreateMultipartPolyline
'    'Get the vertex enumerator
'    pEnumVertex = pPointCollection.EnumVertices
'    'Cocreate a point object for the Query methods
'    'The point doesn't need to be created each time
'    pQueryVertex = New Point
'    'Reset the enum
'    pEnumVertex.Reset()
'    'Get the next vertex
'    pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    Debug.Print("***** Next *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Reset the enum
'    pEnumVertex.Reset()
'    'Query the next vertex - have to cocreate the point
'    'QueryNext is faster than Next, because the method doesn't have
'    'to create the point each time
'    pEnumVertex.QueryNext(pQueryVertex, lOutPart, lOutVertex)
'    Debug.Print("***** QueryNext *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Clone the enumerator
'    pClonedEnumVertex = pEnumVertex.Clone
'    'Get the next vertex
'    pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    Debug.Print("***** Next *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Get the next vertex of the cloned enumerator
'    pClonedEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    Debug.Print("***** Next - Cloned Enum *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Get the previous vertex
'    pEnumVertex.Previous(pOutVertex, lOutPart, lOutVertex)
'    Debug.Print("***** Previous *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    'Query the previous vertex
'    pEnumVertex.QueryPrevious(pQueryVertex, lOutPart, lOutVertex)
'    Debug.Print("***** QueryPrevious *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Set the enumerator at Part Index = 0 and Vertex Index = 4 (Last vertex)
'    pEnumVertex.SetAt(0, 4)
'    Debug.Print("***** IsLastInPart *****")
'    Debug.Print("IsLastInPart : " & pEnumVertex.IsLastInPart)
'    pEnumVertex.NextInPart(pOutVertex, lOutVertex)
'    Debug.Print("***** NextInPart - Last Vertex *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Next in part will return the first vertex of the current part
'    'because the enumerator is located on the last vertex on that part
'    pEnumVertex.NextInPart(pOutVertex, lOutVertex)
'    Debug.Print("***** NextInPart - First Vertex *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'QueryNextInPart
'    pEnumVertex.QueryNextInPart(pOutVertex, lOutVertex)
'    Debug.Print("***** QueryNextInPart *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Get the previous vertex
'    pEnumVertex.ResetToEnd()
'    pEnumVertex.Previous(pOutVertex, lOutPart, lOutVertex)
'    Debug.Print("***** Previous - After ResetToEnd *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Reset the enumerator
'    pEnumVertex.Reset()
'    'Skip two vertices forward
'    pEnumVertex.Skip(2)
'    pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    Debug.Print("***** Skip forward *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Skip two vertices backward
'    pEnumVertex.Skip(-2)
'    pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    Debug.Print("***** Skip backward *****")
'    Debug.Print("Vertex coordinates X, Y : " & pOutVertex.X & ", " & pOutVertex.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    pEnumVertex.WKSNext(OutWKSPoint, lOutPart, lOutVertex)
'    Debug.Print("***** WKSNext *****")
'    Debug.Print("Vertex coordinates X, Y : " & OutWKSPoint.X & ", " & OutWKSPoint.Y & " - Part Index : " & lOutPart & " - Vertex Index : " & lOutVertex)
'    'Make the polyline Ms Aware, Zs Aware, IDsAware
'    'to be able to assign Ms, Zs, Ids to its vertices
'    pMAware = pPointCollection
'    pMAware.MAware = True
'    pZAware = pPointCollection
'    pZAware.ZAware = True
'    pPtIdAware = pPointCollection
'    pPtIdAware.PointIDAware = True
'    pEnumVertex = pPointCollection.EnumVertices
'    'Set the vertex attributes
'    pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    While Not pOutVertex Is Nothing
'        pEnumVertex.put_M(CDbl(i))
'        pEnumVertex.put_Z(CDbl(i))
'        pEnumVertex.put_ID(CDbl(i))
'        i = i + 1
'        pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    End While
'    'Move the first point of the polyline by -10, -10
'    pEnumVertex.Reset()
'    pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    pEnumVertex.put_X(pOutVertex.X - 10)
'    pEnumVertex.put_Y(pOutVertex.Y - 10)

'    'Loop over the vertices and print the values
'    pEnumVertex.Reset()
'    pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    While Not pOutVertex Is Nothing
'        Debug.Print("Vertex coordinates X, Y, M, Z, ID : " & pOutVertex.X & ", " & pOutVertex.Y & ", " & pOutVertex.M & ", " & pOutVertex.Z & ", " & pOutVertex.ID)
'        pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
'    End While


'End Sub