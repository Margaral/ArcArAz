Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

<ComClass(cmd_6_4_VertexAtDistance.ClassId, cmd_6_4_VertexAtDistance.InterfaceId, cmd_6_4_VertexAtDistance.EventsId), _
 ProgId("ArcArAz.cmd_6_4_VertexAtDistance")> _
Public NotInheritable Class cmd_6_4_VertexAtDistance
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "fc1e1139-653f-4215-bca9-39692b23f6b7"
    Public Const InterfaceId As String = "1b4545f9-0383-4947-a6bc-18e4fc664615"
    Public Const EventsId As String = "dabba163-e896-4ba8-b0aa-d14800c5c4a0"
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
        MyBase.m_caption = "Vertices at specific distance"   'localizable text 
        MyBase.m_message = "Create new vertices separated the specific distance"   'localizable text 
        MyBase.m_toolTip = "Create new vertices separated the specific distance" 'localizable text 
        MyBase.m_name = "ArcArAz-Vertex_VerticesAtDistanceCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_6_4_VertexAtDistance.OnClick implementation

        Dim dblPointDist As Double = InputBox("Distance between vertex: ")

        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pInLayer As ILayer = pMxDoc.SelectedLayer

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a POLYGON or POLYLINE feature layer in the TOC", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim pInFLayer As IFeatureLayer
        If TypeOf pInLayer Is IFeatureLayer Then  'check if selected layer is a feature layer
            pInFLayer = pMxDoc.SelectedLayer  'set selected layer as input feature layer
        Else
            MsgBox("Select a POLYGON or POLYLINE feature layer in the TOC", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim pInFClass As IFeatureClass = pInFLayer.FeatureClass

        Dim pFSelection As IFeatureSelection = pInFLayer

        Dim pSelSet As ISelectionSet = pFSelection.SelectionSet

        Dim pInFCursor As IFeatureCursor = Nothing

        If pSelSet.Count <> 0 Then
            'use selected features from input feature class
            pFSelection.SelectionSet.Search(Nothing, True, pInFCursor)
        Else
            'use all features if none are selected
            pInFCursor = pInFClass.Search(Nothing, True)
        End If

        Dim dblLoop As Double

        Dim pInFeature As IFeature
        pInFeature = pInFCursor.NextFeature

        Do While Not pInFeature Is Nothing

            Dim pInGeometry As IGeometry
            pInGeometry = pInFeature.Shape

            Dim pPolyCurve As IPolycurve
            pPolyCurve = pInGeometry

            Dim blnSplitHapp As Boolean
            blnSplitHapp = True
            dblLoop = dblPointDist
            'Do While blnSplitHapp = True
            Do While dblLoop < pPolyCurve.Length
                Dim lngPart As Long
                Dim lngSegment As Long
                On Error Resume Next
                pPolyCurve.SplitAtDistance(dblLoop, False, False, blnSplitHapp, lngPart, lngSegment)
                dblLoop = dblLoop + dblPointDist
                pInFeature.Shape = pPolyCurve
                pInFeature.Store()
            Loop
            pInFeature = pInFCursor.NextFeature

        Loop
        pMxDoc.ActiveView.Refresh()
        MsgBox("Succeeded at " & Now, MsgBoxStyle.Information)
    End Sub
End Class



