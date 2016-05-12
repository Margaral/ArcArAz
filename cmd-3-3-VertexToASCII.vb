Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase
Imports System.IO
Imports ESRI.ArcGIS.esriSystem

<ComClass(cmd3_3_VertexToASCII.ClassId, cmd3_3_VertexToASCII.InterfaceId, cmd3_3_VertexToASCII.EventsId), _
 ProgId("ArcArAz.cmd3_3_VertexToASCII")> _
Public NotInheritable Class cmd3_3_VertexToASCII
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "dbcfec17-8e84-4515-8da1-af44adea10ff"
    Public Const InterfaceId As String = "3f1218bf-a59f-4b4a-8e01-c49cd3779332"
    Public Const EventsId As String = "9104a502-4ebb-4eb7-9d8f-9e4f0bacfdc1"
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
        MyBase.m_category = "ArcArAz-SpatialFields"  'localizable text 
        MyBase.m_caption = "Vertex Coordinates to ASCII"   'localizable text 
        MyBase.m_message = "Create a .txt file within the X, Y and Z coordinates"   'localizable text 
        MyBase.m_toolTip = "Create a .txt file within the X, Y and Z coordinates" 'localizable text 
        MyBase.m_name = "ArcArAz-SpatialFields_VertexToASCIICommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd3_3_VertexToASCII.OnClick implementation

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pInLayer As IFeatureLayer = pMxDoc.SelectedLayer

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a POLYGON or POLYLINE feature class in the TOC ", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim pInFeatClass As IFeatureClass = pInLayer.FeatureClass
        Dim pDataSet As IDataset = pInFeatClass
        Dim pWorkspace As IWorkspace = pDataSet.Workspace
        Dim strTxtWholeName As String = pWorkspace.PathName & "\" & pInLayer.Name & "_SHP2ASCII.txt"

        If File.Exists(strTxtWholeName) Then
            File.Delete(strTxtWholeName)
        End If

        Dim objWriter As StreamWriter = New StreamWriter(strTxtWholeName)

        For i = 0 To pInFeatClass.FeatureCount(Nothing) - 1
            Dim pFeat As IFeature = pInFeatClass.GetFeature(i)
            Dim pPointColl As IPointCollection = pFeat.Shape
            For j = 0 To pPointColl.PointCount - 1
                objWriter.WriteLine(Math.Round(pPointColl.Point(j).X, 2) & " " & Math.Round(pPointColl.Point(j).Y, 2) & " " & Math.Round(pPointColl.Point(j).Z, 3))
            Next
        Next
        objWriter.Close()
        MsgBox("Process finished at " & Now, MsgBoxStyle.Information)
    End Sub
End Class



