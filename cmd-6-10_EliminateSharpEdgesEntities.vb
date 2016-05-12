Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

<ComClass(cmd_6_10_EliminateSharpEdgesEntities.ClassId, cmd_6_10_EliminateSharpEdgesEntities.InterfaceId, cmd_6_10_EliminateSharpEdgesEntities.EventsId), _
 ProgId("ArcArAz.cmd_6_10_EliminateSharpEdgesEntities")> _
Public NotInheritable Class cmd_6_10_EliminateSharpEdgesEntities
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "91b5ce78-cbd4-4055-a97d-99dc5bc7423f"
    Public Const InterfaceId As String = "87cbeef4-d918-4957-ae2f-155ac4ca50d4"
    Public Const EventsId As String = "d9110c38-0b9f-476c-8071-9371e17e0956"
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
        MyBase.m_category = "TRANIN-Vertex"  'localizable text 
        MyBase.m_caption = "Eliminate Sharp Edges from Polygons"   'localizable text 
        MyBase.m_message = "Eliminate sharp edges of individual entities"   'localizable text 
        MyBase.m_toolTip = "Eliminate sharp edges of individual entities" 'localizable text 
        MyBase.m_name = "TRANIN-Vertex-SharpEdgesEntity"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_6_10_EliminateSharpEdgesEntities.OnClick implementation

        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pSel As ISelection = pMap.FeatureSelection
        Dim pMapSel As IEnumFeature = pSel
        pMapSel.Reset()

        Dim limite As Double = InputBox("Set tolerance distance in map units: ")

        Dim pfeat As IFeature = pMapSel.Next()
        Do Until pfeat Is Nothing
            Dim pPoly1 As IPolygon = pfeat.Shape
            Dim pPointColl1 As IPointCollection = pPoly1
            Dim pPointCollCount As Integer = pPointColl1.PointCount - 1
            Dim j As Integer = 0
            While j < pPointCollCount
                On Error Resume Next

                Dim punto1 As IPoint
                Dim punto2 As IPoint
                Dim puntoC As IPoint

                If j = 0 Then
                    punto1 = pPointColl1.Point(pPointColl1.PointCount - 1)
                    punto2 = pPointColl1.Point(1)
                    puntoC = pPointColl1.Point(0)
                ElseIf j < pPointColl1.PointCount - 1 Then
                    punto1 = pPointColl1.Point(j - 1)
                    punto2 = pPointColl1.Point(j + 1)
                    puntoC = pPointColl1.Point(j)
                Else
                    Exit Sub
                End If

                Dim moduloA As Double = Math.Sqrt((puntoC.X - punto1.X) ^ 2 + (puntoC.Y - punto1.Y) ^ 2)
                Dim moduloB As Double = Math.Sqrt((puntoC.X - punto2.X) ^ 2 + (puntoC.Y - punto2.Y) ^ 2)
                Dim prodEscalar As Double = (puntoC.X - punto1.X) * (puntoC.X - punto2.X) + (puntoC.Y - punto1.Y) * (puntoC.Y - punto2.Y)
                Dim alfa As Double = Math.Acos(prodEscalar / (moduloA * moduloB)) * 180 / Math.PI

                If alfa < 60 Then
                    Dim ladoMayor As Double = limite * Math.Sin((180 - alfa) * Math.PI / (2 * 180)) / Math.Sin(alfa * Math.PI / 180)
                    Dim pTable As ITable = pfeat.Table
                    Dim pDataset As IDataset = pTable
                    Dim pWSE As IWorkspace = pDataset.Workspace

                    Dim bordeMenor As Double = Math.Min(moduloA, moduloB)

                    If moduloA > 4 * moduloB Or moduloB > 4 * moduloA Or bordeMenor < ladoMayor Then
                        pWSE.StartEditing(False)
                        pPointColl1.RemovePoints(j, 1)
                        pPointCollCount = pPointColl1.PointCount - 1
                        pfeat.Store()
                        pWSE.StopEditing(True)
                        j = j - 1

                    Else
                        pWSE.StartEditing(False)
                        Dim newPoint1 As IPoint = New ESRI.ArcGIS.Geometry.Point

                        Dim pLine As ILine = New Line
                        pLine.PutCoords(puntoC, punto1)
                        pLine.QueryPoint(esriSegmentExtension.esriNoExtension, ladoMayor, False, newPoint1)
                        pPointColl1.AddPoint(newPoint1, before:=j)

                        Dim newPoint2 As IPoint = New ESRI.ArcGIS.Geometry.Point
                        pLine.PutCoords(puntoC, punto2)
                        pLine.QueryPoint(esriSegmentExtension.esriNoExtension, ladoMayor, False, newPoint2)
                        pPointColl1.AddPoint(newPoint2, after:=j + 1)

                        pPointColl1.RemovePoints(j + 1, 1)
                        pfeat.Store()
                        pWSE.StopEditing(True)
                        j = j + 1
                    End If

                    Dim graphicsContainer As ESRI.ArcGIS.Carto.IGraphicsContainer = CType(pMap, ESRI.ArcGIS.Carto.IGraphicsContainer)

                    'Create a circle with given radius
                    Dim pConstructCArc As IConstructCircularArc = New CircularArc
                    pConstructCArc.ConstructCircle(puntoC, limite * 2, True)

                    'Add the circle segment to a polygon geometry
                    Dim pSegmentCollection As ISegmentCollection = New Polyline
                    pSegmentCollection.AddSegment(pConstructCArc)
                    Dim pPolyline As IPolyline = pSegmentCollection
                    Dim element As ESRI.ArcGIS.Carto.IElement = New LineElement

                    element.Geometry = pPolyline
                    graphicsContainer.AddElement(element, 0)

                Else
                    j = j + 1

                End If

            End While

            GC.Collect()
            pfeat = pMapSel.Next()
        Loop
        pMxDoc.ActiveView.Refresh()
    End Sub
End Class



