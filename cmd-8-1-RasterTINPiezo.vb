Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geometry
Imports System.IO
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesFile


<ComClass(cmd_8_1_RasterTINPiezo.ClassId, cmd_8_1_RasterTINPiezo.InterfaceId, cmd_8_1_RasterTINPiezo.EventsId), _
 ProgId("ArcArAz.cmd_8_1_RasterTINPiezo")> _
Public NotInheritable Class cmd_8_1_RasterTINPiezo
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ac2346ae-ab20-44f7-babc-3610f15e6d45"
    Public Const InterfaceId As String = "6c8396b7-5104-4d6d-91c0-25e5f5c2db96"
    Public Const EventsId As String = "813d384f-b81e-4011-a1cb-077f99897e98"
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
        MyBase.m_category = "ArcArAz-Output"  'localizable text 
        MyBase.m_caption = "Raster / TIN Piezo"   'localizable text 
        MyBase.m_message = "Create a raster or TIN with the piezometry"   'localizable text 
        MyBase.m_toolTip = "Create a raster or TIN with the piezometry" 'localizable text 
        MyBase.m_name = "ArcArAz-Output_RasterTINPiezoCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_8_1_RasterTINPiezo.OnClick implementation
        'If System.IO.Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) & "\VisualGUM\tmp\nodesmesh\") Then
        '    System.IO.Directory.Delete(System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) & "\VisualGUM\tmp\nodesmesh\")
        'End If
        ' Use the OpenFileDialog Class to choose which shapefile to load.

        ' Create a new ShapefileWorkspaceFactory CoClass to create a new workspace
        Dim workspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactoryClass

        ' System.IO.Path.GetDirectoryName(shapefileLocation) returns the directory part of the string. Example: "C:\test\"
        Dim featureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = CType(workspaceFactory.OpenFromFile(folderName, 0), ESRI.ArcGIS.Geodatabase.IFeatureWorkspace) ' Explicit Cast


        Dim templates As ESRI.ArcGIS.Framework.ITemplates = m_application.Templates
        Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra

        Dim openFileDialog As System.Windows.Forms.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        openFileDialog.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) & "\VisualGUM\tmp\"
        openFileDialog.Filter = "ArcArAz Files | *.dat;*.out"      '"Text Files (*.out)|*.out | Data Files  (*.dat)|*.dat"
        openFileDialog.FilterIndex = 2
        openFileDialog.Multiselect = True
        openFileDialog.RestoreDirectory = False
        openFileDialog.Title = "Select files generated by Visual ArcArAz"

        If openFileDialog.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim filesVT() As String = openFileDialog.FileNames()
            Dim intNodes As String = " " 'Trim(Mid(strLine2, 1, 5))
            Dim intElements As String = " "  'Trim(Mid(strLine2, 6, 5))

            Dim strPathGeneral As String = System.IO.Path.GetDirectoryName(filesVT(0).ToString) 'sin la última contrabarra
            Dim strName As String = "nodesmesh" '& Now()
            'MsgBox(strPathGeneral & "\" & strName)

            Dim pMxDoc As IMxDocument = m_application.Document
            Dim pMap As IMap = pMxDoc.ActiveView
            'Dim pLayer As IFeatureLayer = pMxDoc.SelectedLayer
            'If pLayer Is Nothing Then  'Check if no input layer is selected
            '    MsgBox("Select a feature class in the TOC to create a TIN", vbCritical, "Incompatible input layer")
            '    Exit Sub
            'End If
            'Dim pFeatClass As IFeatureClass = pLayer.FeatureClass
            'Dim pDataSet As IDataset = pFeatClass
            'Dim strPath As String = pDataSet.Workspace.PathName 'este string coge también el nombre de la bd con la extensión .mdb
            'Dim strFolder As String = Microsoft.VisualBasic.Left(strPath, InStrRev(strPath, "\")) 'este string acaba en nombrecarpetaBD\
            'Dim pGDS As IGeoDataset = pFeatClass
            Dim pEnv As IEnvelope = New EnvelopeClass()
            pEnv.PutCoords(-10000, -100000, 10000, 100000)
            'pEnv = pGDS.Extent
            'pEnv.SpatialReference = pGDS.SpatialReference
            Dim pTinEdit As ITinEdit = New TinClass()
            pTinEdit.InitNew(pEnv)


            'Open GRI.dat file to get the coordinates of nodes
            Dim objGRIreader As StreamReader = Nothing
            Dim objMHHreader As StreamReader = Nothing
            For Each f As String In filesVT
                Dim strNameDAT As String = System.IO.Path.GetFileNameWithoutExtension(f)
                If strNameDAT.Contains("MSH") = True Then
                    objGRIreader = New StreamReader(f)
                ElseIf strNameDAT.Contains("MHH") = True Then
                    objMHHreader = New StreamReader(f)
                End If
            Next
            For i As Integer = 0 To 6
                objGRIreader.ReadLine()
            Next

            Dim strLineElements As String = objGRIreader.ReadLine()
            intElements = Trim(Mid(strLineElements, 6, 5))
            intNodes = Trim(Mid(strLineElements, 1, 5))

            Dim psbar As IStatusBar = m_application.StatusBar
            Dim pPro As IStepProgressor = psbar.ProgressBar

            pPro.MinRange = 1
            pPro.MaxRange = intNodes + 4
            pPro.StepValue = (intNodes + 4) / 100
            pPro.Step()
            pPro.Show()

            Dim strEmptyLine As String = objGRIreader.ReadLine()
            Dim strEmptyLine1 As String = objMHHreader.ReadLine()

            Dim intNumLine As Integer = 3
            Do Until intNumLine = intNodes + 3
                Dim strGRIline As String = objGRIreader.ReadLine()
                Dim strMHHline As String = objMHHreader.ReadLine()

                Dim pPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
                Dim pZAware As IZAware = pPoint
                pZAware.ZAware = True

                On Error Resume Next
                pPoint.X = CDbl(Trim(Mid(strGRIline, 6, 10)))
                pPoint.Y = CDbl(Trim(Mid(strGRIline, 16, 10)))
                pPoint.Z = CDbl(Trim(Mid(strMHHline, 6, 11)))

                pTinEdit.AddShapeZ(pPoint, esriTinSurfaceType.esriTinMassPoint, Nothing, Nothing)

                intNumLine = intNumLine + 1
                pPro.Position = intNumLine

            Loop
            pPro.Hide()
            pTinEdit.SaveAs(strPathGeneral & "\" & strName)
            pTinEdit.StopEditing(True)
            Dim pTINlayer As ITinLayer = New TinLayer
            pTINlayer.Name = strName
            pTINlayer.Dataset = pTinEdit
            pMap.AddLayer(pTINlayer)

        End If
    End Sub
End Class



