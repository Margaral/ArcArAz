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

<ComClass(cmd_6_5_MaxDistanceVertex.ClassId, cmd_6_5_MaxDistanceVertex.InterfaceId, cmd_6_5_MaxDistanceVertex.EventsId), _
 ProgId("ArcArAz.cmd_6_5_MaxDistanceVertex")> _
Public NotInheritable Class cmd_6_5_MaxDistanceVertex
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "235d2131-5de6-4c08-a5ae-d7805e2d1389"
    Public Const InterfaceId As String = "9383941c-f2c5-4cc2-a3bf-774503227d26"
    Public Const EventsId As String = "962d5e36-edc8-49ba-9d65-d0002f6278e5"
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
        MyBase.m_caption = "Maximum distance between vertices"   'localizable text 
        MyBase.m_message = "Create vertices at a maximum specified distance"   'localizable text 
        MyBase.m_toolTip = "Create vertices at a maximum specified distance" 'localizable text 
        MyBase.m_name = "ArcArAz-Vertex_MaxDistanceVertex"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_6_5_MaxDistanceVertex.OnClick implementation
        Dim dblPointDist As Double = InputBox("Distance between vertex: ")

        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap


        Dim pLayerSet1 As ISet = pMxDoc.CurrentContentsView.SelectedItem
        pLayerSet1.Reset()

        Dim pInLayer As ILayer = pLayerSet1.Next

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a POLYGON or POLYLINE feature layers in the TOC", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        While Not pInLayer Is Nothing

            Dim pInFLayer As IFeatureLayer
            If TypeOf pInLayer Is IFeatureLayer Then  'check if selected layer is a feature layer
                pInFLayer = pInLayer 'pMxDoc.SelectedLayer  'set selected layer as input feature layer
            Else
                pInLayer = pLayerSet1.Next
                Continue While
                'MsgBox("Select a POLYGON or POLYLINE feature layer in the TOC", vbCritical, "Incompatible input layer")
                'Exit Sub
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

            Dim pInFeature As IFeature
            pInFeature = pInFCursor.NextFeature

            Do While Not pInFeature Is Nothing

                Dim pInGeometry As IGeometry = pInFeature.Shape
                Dim pPolyCurve As IPolycurve = pInGeometry
                pPolyCurve.Densify(dblPointDist, 0)
                pInFeature.Shape = pPolyCurve
                pInFeature.Store()
                pInFeature = pInFCursor.NextFeature

            Loop
            pInLayer = pLayerSet1.Next
        End While
        pMxDoc.ActiveView.Refresh()
        MsgBox("Succeeded at " & Now, MsgBoxStyle.Information)
    End Sub
End Class



