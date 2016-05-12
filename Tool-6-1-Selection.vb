Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase


<ComClass(Tool_6_1_VertexCount.ClassId, Tool_6_1_VertexCount.InterfaceId, Tool_6_1_VertexCount.EventsId), _
 ProgId("ArcArAz.Tool_6_1_VertexCount")> _
Public NotInheritable Class Tool_6_1_VertexCount
    Inherits BaseTool

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "054ca18d-d0a0-421b-8aeb-38a5b716fcfd"
    Public Const InterfaceId As String = "3806254b-7ee8-44dd-8009-215c856a4e76"
    Public Const EventsId As String = "771388b9-f1d7-48fb-a065-c92e180df2ed"
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
    Private m_Document As IDocument
    Private m_CommandBars As ICommandBars

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()

        ' TODO: Define values for the public properties
        ' TODO: Define values for the public properties
        MyBase.m_category = "CSIC-IDAEA"  'localizable text 
        MyBase.m_caption = "Select entity to submit"   'localizable text 
        MyBase.m_message = "Double click on feature of interest"   'localizable text 
        MyBase.m_toolTip = "Double click on feature of interest" 'localizable text 
        MyBase.m_name = "CSIC-IDAEA_SelectionTool"  'unique id, non-localizable (e.g. "MyCategory_ArcMapTool")

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

    Public Overrides Sub OnClick()
        'TODO: Add Tool_6_1_VertexCount.OnClick implementation
        'TODO: Add SelectFeatureTool.OnClick implementation
        Dim m_UID As New UID
        Dim m_CommandItem As ICommandItem

        'locate the real Select Elements tool and call it
        m_UID.Value = "{C22579D1-BC17-11D0-8667-0000F8751720}"
        m_Document = m_application.Document
        m_CommandBars = m_Document.CommandBars
        m_CommandItem = m_CommandBars.Find(m_UID)
        m_CommandItem.Execute()
        'done

        'locate the real Select Features tool and call it
        m_UID.Value = "{78FFF793-98B4-11D1-873B-0000F8751720}"
        m_Document = m_application.Document
        m_CommandBars = m_Document.CommandBars
        m_CommandItem = m_CommandBars.Find(m_UID)
        m_CommandItem.Execute()
        'done
    End Sub

    Public Overrides Sub OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add Tool_6_1_VertexCount.OnMouseDown implementation
        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pPointColl As IPointCollection
        Dim pGraphics As IGraphicsContainer = pMxDoc.FocusMap
        Dim pActiveView As IActiveView = pGraphics

        Dim i As Integer
        For i = 0 To pMap.LayerCount - 1

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
                    MsgBox(pPointColl.PointCount)
                    pFeat = pInFCursor.NextFeature
                Next
            End If
        Next


    End Sub

    Public Overrides Sub OnMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add Tool_6_1_VertexCount.OnMouseMove implementation
    End Sub

    Public Overrides Sub OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
        'TODO: Add Tool_6_1_VertexCount.OnMouseUp implementation
    End Sub
End Class

