Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesGDB
Imports System.IO
Imports ESRI.ArcGIS.Geometry

<ComClass(cmd_8_7_parametros.ClassId, cmd_8_7_parametros.InterfaceId, cmd_8_7_parametros.EventsId), _
 ProgId("ArcArAz.cmd_8_7_parametros")> _
Public NotInheritable Class cmd_8_7_parametros
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "642e8a59-538f-46ee-9d15-2a6ff9c0d9d0"
    Public Const InterfaceId As String = "356cc935-8b8a-456c-b60c-77a8257bd333"
    Public Const EventsId As String = "6c48db54-ff2e-4e3b-8221-cedc255dc6ef"
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
        MyBase.m_caption = "Parameters Nodes"   'localizable text 
        MyBase.m_message = "Parameters Nodes"   'localizable text 
        MyBase.m_toolTip = "Parameters Nodes" 'localizable text 
        MyBase.m_name = "ArcArAz-Output_ParametersNodes"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_8_7_parametros.OnClick implementation

        ' Use the OpenFileDialog Class to choose VT files.

        ' Create a new ShapefileWorkspaceFactory CoClass to create a new workspace
        Dim workspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactoryClass

        ' System.IO.Path.GetDirectoryName(shapefileLocation) returns the directory part of the string. Example: "C:\test\"
        Dim featureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = CType(workspaceFactory.OpenFromFile(folderName, 0), ESRI.ArcGIS.Geodatabase.IFeatureWorkspace) ' Explicit Cast


        Dim templates As ESRI.ArcGIS.Framework.ITemplates = m_application.Templates
        Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra

        Dim openFileDialog As System.Windows.Forms.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        'openFileDialog.InitialDirectory = "c:\"
        openFileDialog.Filter = "TD2VTK Files | *.vtk"      '"Text Files (*.out)|*.out | Data Files  (*.dat)|*.dat"
        openFileDialog.FilterIndex = 2
        openFileDialog.Multiselect = True
        openFileDialog.RestoreDirectory = False
        openFileDialog.Title = "Select files generated by DT2VTK"

        If openFileDialog.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim filesVT() As String = openFileDialog.FileNames()
            Dim intNodes As String = " " 'Trim(Mid(strLine2, 1, 5))
            Dim intElements As String = " "  'Trim(Mid(strLine2, 6, 5))

            '''''''
            '''''''Create a feature class to store nodes as points
            '''''''

            Dim strPathGeneral As String = System.IO.Path.GetDirectoryName(filesVT(0).ToString) 'sin la última contrabarra
            Dim strName As String = "NodesMesh" '& Now()

            Const strShapeFieldName As String = "Shape"

            ' Open the folder to contain the shapefile as a workspace

            Dim pWorkspaceFactory As IWorkspaceFactory = New ShapefileWorkspaceFactory
            Dim pFWS As IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(strPathGeneral, 0)

            ' Delete an old feature class
            If System.IO.File.Exists(strPathGeneral & "\" & strName & ".shp") Then
                Dim pDelFC As IFeatureClass = pFWS.OpenFeatureClass(strName)
                Dim pDataset As IDataset = pDelFC
                pDataset.Delete()
            End If

            ' Set up a simple fields collection
            Dim pFields As IFields
            Dim pFieldsEdit As IFieldsEdit
            pFields = New Fields
            pFieldsEdit = pFields

            Dim pField As IField
            Dim pFieldEdit As IFieldEdit

            ' Make the shape field
            ' it will need a geometry definition, with a spatial reference
            pField = New Field
            pFieldEdit = pField
            pFieldEdit.Name_2 = strShapeFieldName
            pFieldEdit.Type_2 = esriFieldType.esriFieldTypeGeometry

            Dim pGeomDef As IGeometryDef
            Dim pGeomDefEdit As IGeometryDefEdit
            pGeomDef = New GeometryDef
            pGeomDefEdit = pGeomDef
            With pGeomDefEdit
                .GeometryType_2 = esriGeometryType.esriGeometryPoint
                .HasZ_2 = True
                .SpatialReference_2 = New UnknownCoordinateSystem 'pInFeatClass.GetFeature(0).Shape.SpatialReference
            End With
            pFieldEdit.GeometryDef_2 = pGeomDef
            pFieldsEdit.AddField(pField)

            ' Add another miscellaneous text field
            pField = New Field
            pFieldEdit = pField
            With pFieldEdit
                .Length_2 = 10
                .Name_2 = "MHH"
                .Type_2 = esriFieldType.esriFieldTypeDouble
            End With
            pFieldsEdit.AddField(pField)

            ' Create the shapefile
            ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
            Dim pFeatClass As IFeatureClass = pFWS.CreateFeatureClass(strName, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, strShapeFieldName, "")

            'Add shapefile to TOC
            'Create a new ShapefileWorkspaceFactory object and open a shapefile folder

            Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(strPathGeneral, 0)

            'Create a new FeatureLayer and assign a shapefile to it
            Dim pFeatureLayer As IFeatureLayer = New FeatureLayer
            pFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(strName)
            pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName

            'Add the FeatureLayer to the focus map
            Dim pMxDoc As IMxDocument = m_application.Document
            Dim pMap As IMap = pMxDoc.FocusMap
            pMap.AddLayer(pFeatureLayer)
            pMap.MoveLayer(pFeatureLayer, (pMxDoc.FocusMap.LayerCount))
            ''''''''
            ''''''''
            ''''''''


            'Open GRI.dat file to get the coordinates of nodes
            Dim objGRIreader As StreamReader = Nothing
            Dim objReader2 As StreamReader = Nothing
            For Each f As String In filesVT
                Dim strNameDAT As String = System.IO.Path.GetFileNameWithoutExtension(f)
                If strNameDAT.Contains("HH") = True Then
                    objGRIreader = New StreamReader(f)
                    objReader2 = New StreamReader(f)
                End If
            Next

            Dim psbar As IStatusBar = m_application.StatusBar
            Dim pPro As IStepProgressor = psbar.ProgressBar

            For i = 0 To 6
                Dim LineWithNumber As String = objGRIreader.ReadLine()
                If i = 5 Then
                    intNodes = CDbl(Trim(Mid(LineWithNumber, 7, 8)))
                End If
            Next

            For i = 0 To intNodes + 10
                Dim LineZone As String = objReader2.ReadLine()
                If i = intNodes + 8 Then
                    'MsgBox(LineZone)
                    intElements = CDbl(Trim(Mid(LineZone, 6, 6)))
                End If
            Next

            For i = 0 To 2 * intElements + 3
                Dim LineZone As String = objReader2.ReadLine()
            Next

            pPro.MinRange = 1
            pPro.MaxRange = intNodes
            pPro.StepValue = intNodes / 100
            pPro.Step()
            pPro.Show()

            Dim intNumLine As Integer = 6
            Do Until intNumLine = intNodes + 6
                Dim strPORline As String = objGRIreader.ReadLine()
                Dim strLine2 As String = objReader2.ReadLine()

                Dim pPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
                Dim pZAware As IZAware = pPoint
                pZAware.ZAware = True

                pPoint.X = CDbl(Trim(Mid(strPORline, 1, 15)))
                pPoint.Y = CDbl(Trim(Mid(strPORline, 16, 16)))
                pPoint.Z = CDbl(Trim(Mid(strPORline, 33, 16)))

                Dim pOutFeat As IFeature = pFeatClass.CreateFeature
                pOutFeat.Shape = pPoint
                pOutFeat.Value(2) = CDbl(Trim(strLine2)) 'intNumLine - 6
                'pOutFeat.Value(3) = CDbl(Trim(strLine2))
                pOutFeat.Store()

                intNumLine = intNumLine + 1
                pPro.Position = intNumLine
            Loop
            pPro.Hide()
            pMxDoc.ActiveView.ContentsChanged()
            pMxDoc.UpdateContents()

            '*************************************************


        End If


    End Sub
End Class



