Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.DataSourcesGDB
Imports System.IO

<ComClass(cmd_8_4_SHPNodeMesh.ClassId, cmd_8_4_SHPNodeMesh.InterfaceId, cmd_8_4_SHPNodeMesh.EventsId), _
 ProgId("ArcArAz.cmd_8_4_SHPNodeMesh")> _
Public NotInheritable Class cmd_8_4_SHPNodeMesh
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "b95e349b-81cd-4701-b320-88eab41a9d7f"
    Public Const InterfaceId As String = "8960779e-4485-4fc8-ae44-93ba44e2f2e1"
    Public Const EventsId As String = "dd0ccc35-e40a-4984-8333-dd500345347f"
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
        MyBase.m_caption = "Shapefile with mesh nodes"   'localizable text 
        MyBase.m_message = "Create a point shapefile with the mesh nodes"   'localizable text 
        MyBase.m_toolTip = "Create a point shapefile with the mesh nodes" 'localizable text 
        MyBase.m_name = "ArcArAz-Output_SHPMeshNodesCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_8_4_SHPNodeMesh.OnClick implementation

        ' Use the OpenFileDialog Class to choose VT files.
        Dim templates As ESRI.ArcGIS.Framework.ITemplates = m_application.Templates
        Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra

        ' Create a new ShapefileWorkspaceFactory CoClass to create a new workspace
        Dim workspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactoryClass

        ' System.IO.Path.GetDirectoryName(shapefileLocation) returns the directory part of the string. Example: "C:\test\"
        Dim featureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = CType(workspaceFactory.OpenFromFile(folderName, 0), ESRI.ArcGIS.Geodatabase.IFeatureWorkspace) ' Explicit Cast


        Dim openFileDialog As System.Windows.Forms.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        openFileDialog.InitialDirectory = "c:\"
        openFileDialog.Filter = "ArcArAz Files | *.dat;*.out"      '"Text Files (*.out)|*.out | Data Files  (*.dat)|*.dat"
        openFileDialog.FilterIndex = 2
        openFileDialog.Multiselect = True
        openFileDialog.RestoreDirectory = False
        openFileDialog.Title = "Select files generated by Visual ArcArAz"

        If openFileDialog.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim filesVT() As String = openFileDialog.FileNames()

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
                .Name_2 = "Node"
                .Type_2 = esriFieldType.esriFieldTypeDouble
            End With
            pFieldsEdit.AddField(pField)

            ' Add another miscellaneous text field
            pField = New Field
            pFieldEdit = pField
            With pFieldEdit
                .Length_2 = 10
                .Name_2 = "Height"
                .Type_2 = esriFieldType.esriFieldTypeDouble
            End With
            pFieldsEdit.AddField(pField)

            ' Add another miscellaneous text field
            pField = New Field
            pFieldEdit = pField
            With pFieldEdit
                .Length_2 = 10
                .Name_2 = "Head"
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
            Dim pInsertFeatureBuffer As IFeatureBuffer = pFeatClass.CreateFeatureBuffer
            Dim pInsertFeatureCursor As IFeatureCursor = pFeatClass.Insert(True)
            Dim NewFeatureCount As Integer = 0


            'Open GRI.dat file to get the coordinates of nodes
            Dim objGRIreader As StreamReader = Nothing
            Dim objMHHreader As StreamReader = Nothing
            Dim objDIMreader As StreamReader = Nothing
            For Each f As String In filesVT
                Dim strNameDAT As String = System.IO.Path.GetFileNameWithoutExtension(f)
                If strNameDAT.Contains("GRI") Then '= True Or strNameDAT.Contains("gri1") = True Then
                    objGRIreader = New StreamReader(f)
                ElseIf strNameDAT.Contains("MHH") = True Then
                    objMHHreader = New StreamReader(f)
                ElseIf strNameDAT.Contains("DIM") = True Then
                    objDIMreader = New StreamReader(f)
                End If
            Next

            Dim intNodes As String = " " 'Trim(Mid(strLine2, 1, 5))
            Dim intElements As String = " "  'Trim(Mid(strLine2, 6, 5))

            Dim strNodesElements As String = Nothing
            For i As Integer = 0 To 6
                strNodesElements = objDIMreader.ReadLine()
            Next
            strNodesElements = objDIMreader.ReadLine()
            intElements = Trim(Mid(strNodesElements, 1, 5))
            intNodes = Trim(Mid(strNodesElements, 6, 5))
            MsgBox(intNodes & " nodes; " & intElements & " elements")

            'Dim strLineElements As String = objGRIreader.ReadLine()

            'Dim strLineNodes As String = objGRIreader.ReadLine()
            'intNodes = Trim(Mid(strNodesElements, 10, 6))

            Dim psbar As IStatusBar = m_application.StatusBar
            Dim pPro As IStepProgressor = psbar.ProgressBar

            pPro.MinRange = 1
            pPro.MaxRange = intNodes
            pPro.StepValue = intNodes / 100
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

                pPoint.X = CDbl(Trim(Mid(strGRIline, 6, 10)))
                pPoint.Y = CDbl(Trim(Mid(strGRIline, 16, 10)))
                pPoint.Z = CDbl(Trim(Mid(strGRIline, 26, 10)))

                pInsertFeatureBuffer.Value(2) = intNumLine - 2
                pInsertFeatureBuffer.Value(3) = pPoint.Z
                pInsertFeatureBuffer.Value(4) = CDbl(Trim(Mid(strMHHline, 6, 15)))
                pInsertFeatureBuffer.Shape = pPoint 'pFeature.Shape
                pInsertFeatureCursor.InsertFeature(pInsertFeatureBuffer)


                NewFeatureCount = NewFeatureCount + 1
                'Flush the feature cursor every 100 features
                'This is safer because you can write code to handle a flush error
                'If you don't flush the feature cursor it will automatically flush but
                'after all of your code executes at which time you have no control
                If NewFeatureCount = 100 Then
                    pInsertFeatureCursor.Flush()
                    NewFeatureCount = 0
                End If

                intNumLine = intNumLine + 1
                pPro.Position = intNumLine
            Loop
            pPro.Hide()
            pMxDoc.ActiveView.ContentsChanged()
            pMxDoc.UpdateContents()
        End If

    End Sub

End Class