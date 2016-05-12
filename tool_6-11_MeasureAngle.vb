Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geometry

<ComClass(tool_6_11_MeasureAngle.ClassId, tool_6_11_MeasureAngle.InterfaceId, tool_6_11_MeasureAngle.EventsId), _
 ProgId("ArcArAz.tool_6_11_MeasureAngle")> _
Public NotInheritable Class tool_6_11_MeasureAngle
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "658e39fc-1864-43b1-ac8d-3b5163e795ad"
    Public Const InterfaceId As String = "d7deb1df-2a7e-4e02-a520-4fd8b813f82a"
    Public Const EventsId As String = "b0b686c0-adc3-434a-a21a-0bb36c374883"
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
        MyBase.m_category = "ArcArAz-Vertex"  'localizable text 
        MyBase.m_caption = "Measure angle"   'localizable text 
        MyBase.m_message = "Click over map three times: second one for the vertex angle, firt and third for angle sides"   'localizable text 
        MyBase.m_toolTip = "Measure angle" 'localizable text 
        MyBase.m_name = "ArcArAz-Vertex-MeasureAngle"  'unique id, non-localizable (e.g. "MyCategory_ArcMapTool")

        'Try
        '    'TODO: change resource name if necessary
        '    Dim bitmapResourceName As String = Me.GetType().Name + ".bmp"
        '    MyBase.m_bitmap = New Bitmap(Me.GetType(), bitmapResourceName)
        '    MyBase.m_cursor = New System.Windows.Forms.Cursor(Me.GetType(), Me.GetType().Name + ".cur")
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

    Public p1Point As IPoint
    Public p2Point As IPoint
    Public pCPoint As IPoint
    Public counter As Integer = 0

    Public Overrides Sub OnClick()
        p1Point = New ESRI.ArcGIS.Geometry.Point
        p2Point = New ESRI.ArcGIS.Geometry.Point
        pCPoint = New ESRI.ArcGIS.Geometry.Point
        'TODO: Add tool_6_11_MeasureAngle.OnClick implementation
        'Dim pMouseCursor As IMouseCursor = New MouseCursor
        'pMouseCursor.SetCursor(3)
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add tool_6_11_MeasureAngle.OnMouseDown implementation

        If counter = 0 Then
            p1Point.X = X
            p1Point.Y = Y
            counter = 1

        ElseIf counter = 1 Then
            pCPoint.X = X
            pCPoint.Y = Y
            counter = 2

        ElseIf counter = 2 Then
            p2Point.X = X
            p2Point.Y = Y
            counter = 0

            Dim angulo As Double = GetAngulo(pCPoint, p1Point, p2Point)
            MsgBox(angulo)
        End If
    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add tool_6_11_MeasureAngle.OnMouseMove implementation
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add tool_6_11_MeasureAngle.OnMouseUp implementation
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

End Class

