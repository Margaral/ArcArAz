Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase

<ComClass(cmd_6_2_VertexCount.ClassId, cmd_6_2_VertexCount.InterfaceId, cmd_6_2_VertexCount.EventsId), _
 ProgId("ArcArAz.cmd_6_2_VertexCount")> _
Public NotInheritable Class cmd_6_2_VertexCount
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "593fb203-f05f-4bc2-a474-6d925d7e8233"
    Public Const InterfaceId As String = "7bb91907-fe0f-42d0-aa8a-a2978293dddc"
    Public Const EventsId As String = "4c66733a-bbee-42b1-9dfa-7b6aed520b24"
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
        MyBase.m_caption = "Vertex Count"   'localizable text 
        MyBase.m_message = "Number of vertex of selected feature"   'localizable text 
        MyBase.m_toolTip = "Number of vertex of selected feature" 'localizable text 
        MyBase.m_name = "ArcArAz-Vertex_VertexCountCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_6_2_VertexCount.OnClick implementation
        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pPointColl As IPointCollection
        Dim pGraphics As IGraphicsContainer = pMxDoc.FocusMap
        Dim pActiveView As IActiveView = pGraphics

        Dim i As Integer

        For i = 0 To pMap.LayerCount - 1

            If TypeOf pMap.Layer(i) Is IFeatureLayer Then
                Dim pFeatSel As IFeatureSelection = pMap.Layer(i)
                Dim pSelSet As ISelectionSet = pFeatSel.SelectionSet
                Dim pInFCursor As IFeatureCursor = Nothing

                If pSelSet.Count <> 0 Then
                    pFeatSel.SelectionSet.Search(Nothing, True, pInFCursor)
                    Dim pFeat As IFeature
                    pFeat = pInFCursor.NextFeature

                    Dim j As Integer
                    For j = 0 To pSelSet.Count - 1
                        pPointColl = pFeat.Shape
                        MsgBox("Total of vertex in feature " & pFeat.OID & vbCrLf & _
                                "from feature class " & pMap.Layer(i).Name.ToString & ": " & vbCrLf & _
                                pPointColl.PointCount & " vertex", MsgBoxStyle.Information, "Total vertex of selected feature")
                        pFeat = pInFCursor.NextFeature

                    Next
                End If
            End If
        Next


    End Sub
End Class



