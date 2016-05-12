Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Analyst3D

<ComClass(tool_8_1_ParticleTrack.ClassId, tool_8_1_ParticleTrack.InterfaceId, tool_8_1_ParticleTrack.EventsId), _
 ProgId("ArcArAz.tool_8_1_ParticleTrack")> _
Public NotInheritable Class tool_8_1_ParticleTrack
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "bae94db5-6b52-4824-9e17-3c2b7db668c7"
    Public Const InterfaceId As String = "09494e03-789d-47a8-a259-061565eee683"
    Public Const EventsId As String = "f30cb26c-967d-46e9-9299-1e2417e9f445"
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
        ' TODO: Define values for the public properties
        MyBase.m_category = "ArcArAz-Output"  'localizable text 
        MyBase.m_caption = "Particle Track"   'localizable text 
        MyBase.m_message = "Particle Track"   'localizable text 
        MyBase.m_toolTip = "Select on TOC piezometric surface raster. Particle Track" 'localizable text 
        MyBase.m_name = "ArcArAz-Output_ParticleTrack"  'unique id, non-localizable (e.g. "MyCategory_ArcMapTool")

        'Try
        '    'TODO: change resource name if necessary
        '    Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
        '    'MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        '    MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), Me.GetType().Name + ".cur")
        'Catch ex As Exception
        '    System.Diagnostics.Trace.WriteLine(ex.Message, "Invalid Bitmap")
        'End Try
    End Sub

    Private Function PolygonToPolyline(ByRef pPolygon As IPolygon) As IGeometryCollection
        '
        ' Create a new Polyline geometry.
        '
        PolygonToPolyline = New Polyline
        '
        ' Clone the incoming Polygon, so we can use the references to this new Polygons segments.
        '
        Dim pGeoms_Polygon As IGeometryCollection, pClone As IClone
        pClone = pPolygon
        pGeoms_Polygon = pClone.Clone
        '
        ' Transfer the segments from the Polygon Rings to new Paths, and add the Paths
        ' to the new Polyline.
        '
        Dim i As Long, pSegs_Path As ISegmentCollection
        For i = 0 To pGeoms_Polygon.GeometryCount - 1
            pSegs_Path = New Path
            pSegs_Path.AddSegmentCollection(pGeoms_Polygon.Geometry(i))
            PolygonToPolyline.AddGeometry(pSegs_Path)
        Next i
    End Function

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
        'TODO: Add tool_8_1_ParticleTrack.OnClick implementation
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add tool_8_1_ParticleTrack.OnMouseDown implementation
    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add tool_8_1_ParticleTrack.OnMouseMove implementation
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add tool_8_1_ParticleTrack.OnMouseUp implementation
        'TODO: Add TL_BHElocation.OnMouseUp implementation
        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pGraphics As IGraphicsContainer = pMxDoc.FocusMap
        Dim pActiveView As IActiveView = pGraphics

        Dim pRaster As IRaster = Nothing
        If TypeOf pMxDoc.SelectedLayer Is IRasterLayer Then
            Dim pRasterLayer As IRasterLayer = pMxDoc.SelectedLayer
            pRaster = pRasterLayer.Raster
        End If

        Dim screenDisplay As ESRI.ArcGIS.Display.IScreenDisplay = pActiveView.ScreenDisplay
        Dim displayTransformation As ESRI.ArcGIS.Display.IDisplayTransformation = screenDisplay.DisplayTransformation
        Dim pPoint As ESRI.ArcGIS.Geometry.IPoint = displayTransformation.ToMapPoint(X, Y)

        pGraphics.Reset()
        Dim pGraphic As IElement = pGraphics.Next

        Dim blnSplitHapp As Boolean
        Dim dblLoop As Double
        Dim dblPointDist As Double = 10


        Do Until pGraphic Is Nothing
            If pGraphic.HitTest(pPoint.X, pPoint.Y, 0.01) Then
                Dim pRasterSurface As IRasterSurface = New RasterSurface
                pRasterSurface.PutRaster(pRaster, 0)
                Dim pSurface As ISurface = pRasterSurface
                Dim rgbColor As IRgbColor = New ESRI.ArcGIS.Display.RgbColor()
                rgbColor.Red = 255

                Dim pPolyCurve As IPolycurve = pGraphic.Geometry

                blnSplitHapp = True
                dblLoop = dblPointDist
                'Do While blnSplitHapp = True
                Do While dblLoop < pPolyCurve.Length
                    Dim lngPart As Long
                    Dim lngSegment As Long = 10
                    On Error Resume Next
                    pPolyCurve.SplitAtDistance(dblLoop, False, False, blnSplitHapp, lngPart, lngSegment)
                    dblLoop = dblLoop + dblPointDist
                    pGraphic.Geometry = pPolyCurve
                Loop

                Dim pPointColl As IPointCollection = pPolyCurve
                Dim pEnumVertex As IEnumVertex2 = pPointColl.EnumVertices
                Dim pOutVertex As IPoint = Nothing
                Dim lOutPart As Long, lOutVertex As Long

                pEnumVertex.Reset()
                pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
                While Not pOutVertex Is Nothing
                    Dim pPolyline As IPolyline = pSurface.GetSteepestPath(pOutVertex)
                    Dim element As IElement = Nothing
                    '  Line elements
                    Dim simpleLineSymbol As ISimpleLineSymbol = New SimpleLineSymbolClass()
                    simpleLineSymbol.Color = rgbColor
                    simpleLineSymbol.Style = esriSimpleLineStyle.esriSLSSolid
                    simpleLineSymbol.Width = 5

                    Dim lineElement As ILineElement = New LineElementClass()
                    lineElement.Symbol = simpleLineSymbol
                    element = CType(lineElement, IElement) ' Explicit Cast
                    pEnumVertex.Next(pOutVertex, lOutPart, lOutVertex)
                    element.Geometry = pPolyline
                    pGraphics.AddElement(element, 0)

                End While

            End If
            pGraphic = pGraphics.Next
        Loop
        pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
    End Sub
End Class

