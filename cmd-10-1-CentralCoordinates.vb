Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports System.Windows.Forms
Imports ESRI.ArcGIS.DataSourcesFile

<ComClass(cmd_10_1_CentralCoordinates.ClassId, cmd_10_1_CentralCoordinates.InterfaceId, cmd_10_1_CentralCoordinates.EventsId), _
 ProgId("ArcArAz.cmd_10_1_CentralCoordinates")> _
Public NotInheritable Class cmd_10_1_CentralCoordinates
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "3728685e-5b64-4309-af19-de18a32e478b"
    Public Const InterfaceId As String = "132033ea-42e1-425d-abbf-be9ed90bc574"
    Public Const EventsId As String = "0193c51f-db60-4acf-9fcc-b676e352bfcd"
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
        MyBase.m_category = "ArcArAz-ToVT"  'localizable text 
        MyBase.m_caption = "Central Coordinates"   'localizable text 
        MyBase.m_message = "Displace coordinates to 0, 0"   'localizable text 
        MyBase.m_toolTip = "Displace coordinates to 0, 0" 'localizable text 
        MyBase.m_name = "Displace coordinates to 0, 0"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_10_1_CentralCoordinates.OnClick implementation

        'Dim templates As ESRI.ArcGIS.Framework.ITemplates = m_application.Templates
        'Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        'folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra

        'Dim createFolder As New FolderBrowserDialog
        'createFolder.Description = "Select Folder"
        ''createFolder.RootFolder = folderName
        'createFolder.ShowNewFolderButton = True

        'Dim strMovePath As String = Nothing
        'Dim dlgResult As DialogResult = createFolder.ShowDialog()
        'If dlgResult = Windows.Forms.DialogResult.OK Then
        '    strMovePath = createFolder.SelectedPath ' no lleva la contrabarra final
        '    strMovePath = strMovePath & "\"
        'End If

        'Dim shps As New IO.DirectoryInfo(strMovePath)
        'Dim shpsnames As IO.FileInfo() = shps.GetFiles()
        'Dim strNames As IO.FileInfo
        'Dim nombres As ArrayList = New ArrayList
        'For Each strNames In shpsnames
        '    If strNames.ToString.EndsWith(".shp") Then
        '        nombres.Add(strNames.ToString)
        '    End If
        'Next

        'Dim pWSF As IWorkspaceFactory = New ShapefileWorkspaceFactory
        'Dim pFeatWS As IFeatureWorkspace
        'Dim pFeatDS As IFeatureClass

        'If pWSF.IsWorkspace(strMovePath) Then
        '    pFeatWS = pWSF.OpenFromFile(strMovePath, 0)
        '    On Error Resume Next
        '    For s = 0 To nombres.Count - 1
        '        pFeatDS = pFeatWS.OpenFeatureClass(nombres.Item(s).ToString)
        '        MsgBox(nombres.Item(s).ToString)
        '    Next
        'End If




        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim ppSet As ISet = Nothing
        Dim pLayer As ILayer = Nothing
        Dim pLayers As IArray = New Array

        Dim pMxTOC As IContentsView3 = pMxDoc.CurrentContentsView
        Dim p As Object = pMxTOC.SelectedItem
        Dim xMax As Double = 0.0
        Dim xMin As Double = 10000000
        Dim yMax As Double = 0.0
        Dim yMin As Double = 10000000

        If TypeOf p Is ISet Then
            ppSet = p
            ppSet.Reset()
            For i = 0 To ppSet.Count
                pLayer = ppSet.Next
                If Not pLayer Is Nothing Then
                    pLayers.Add(pLayer)
                    If xMax < pLayer.AreaOfInterest.XMax Then
                        xMax = pLayer.AreaOfInterest.XMax
                    End If
                    If xMin > pLayer.AreaOfInterest.XMin Then
                        xMin = pLayer.AreaOfInterest.XMin
                    End If
                    If yMax < pLayer.AreaOfInterest.YMax Then
                        yMax = pLayer.AreaOfInterest.YMax
                    End If
                    If yMin > pLayer.AreaOfInterest.YMin Then
                        yMin = pLayer.AreaOfInterest.YMin
                    End If
                End If
            Next
        End If

        If pLayers.Count = 0 Then  'Check if no input layer is selected
            MsgBox("Select feature classes in TOC to create a safety copy", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim xRes As Double = Microsoft.VisualBasic.Fix(xMin + (xMax - xMin) / 2)
        Dim yRes As Double = Microsoft.VisualBasic.Fix(yMin + (yMax - yMin) / 2)

        If TypeOf p Is ISet Then
            ppSet = p
            ppSet.Reset()

            For i = 0 To ppSet.Count - 1
                pLayer = ppSet.Next
                If Not pLayer Is Nothing Then
                    Dim pFeatLayer As IFeatureLayer = pLayer
                    Dim pFeatClass As IFeatureClass = pFeatLayer.FeatureClass
                    For j = 0 To pFeatClass.FeatureCount(Nothing) - 1
                        Dim pFeat As IFeature = pFeatClass.GetFeature(j)
                        Dim pGeom As IGeometry = pFeat.Shape
                        Dim pTrans2D As ITransform2D = pGeom
                        pTrans2D.Move(-xRes, -yRes)
                        pFeat.Store()
                    Next
                End If
            Next
        End If

        MsgBox("X = -" & xRes & vbCrLf & "Y = -" & yRes, MsgBoxStyle.Information)

        pMxDoc.ActiveView.Extent = pLayer.AreaOfInterest
        pMxDoc.ActiveView.Refresh()
    End Sub
    '    Private Function GetDocLayers(ByVal pMap As IMap, Optional ByVal bOnlySelected As Boolean = True) As IArray

    '        Dim pTOC As IContentsView = Nothing
    '        Dim i As Integer
    '        Dim ppSet As ISet
    '        Dim p
    '        Dim pLayers As IArray
    '        Dim pLayer As ILayer

    '        On Error GoTo GetDocLayers_ERR

    '        GetDocLayers = New Array

    '        If Not bOnlySelected Then
    '            pLayers = New Array
    '            For i = 0 To pMap.LayerCount - 1
    '                pLayers.Add(pMap.Layer(i))
    '            Next
    '            GetDocLayers = pLayers
    '            Exit Function
    '        Else
    '            Dim pMxTOC As IContentsView
    '            pMxTOC = pMap.ContentsView(0)
    '        End If

    '        If Not pTOC Is Nothing Then
    '            If (pTOC.SelectedItem) Is Nothing Then Exit Function
    '            p = pTOC.SelectedItem
    '        ElseIf Not pMap Is Nothing Then
    '            If pMap.SelectedItem Is Nothing Then Exit Function
    '            p = pMap.SelectedItem
    '        End If

    '        pLayers = New Array

    '        If TypeOf p Is ISet Then
    '            ppSet = p
    '            ppSet.Reset()
    '            For i = 0 To ppSet.Count
    '                pLayer = ppSet.Next
    '                If Not pLayer Is Nothing Then
    '                    pLayers.Add(pLayer)
    '                End If
    '            Next
    '        ElseIf TypeOf p Is ILayer Then
    '            pLayer = p
    '            pLayers.Add(pLayer)
    '        End If

    '        GetDocLayers = pLayers

    '        Exit Function

    'GetDocLayers_ERR:
    '        Debug.Print("GetDocLayers_ERR: " & Err.Description)
    '        Debug.Assert(0)

    '    End Function
End Class



