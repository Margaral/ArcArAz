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

<ComClass(cmd_8_6_POR.ClassId, cmd_8_6_POR.InterfaceId, cmd_8_6_POR.EventsId), _
 ProgId("ArcArAz.cmd_8_6_POR")> _
Public NotInheritable Class cmd_8_6_POR
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "83c602e8-28a4-4683-a35e-7ba99bd00466"
    Public Const InterfaceId As String = "ff0d6680-3314-4c7c-81be-f369185f4e3b"
    Public Const EventsId As String = "234cda7b-da45-4c3f-b582-6a441844d022"
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
        MyBase.m_caption = "Parameters Elements"   'localizable text 
        MyBase.m_message = "Parameters Elements"   'localizable text 
        MyBase.m_toolTip = "Parameters Elements" 'localizable text 
        MyBase.m_name = "ArcArAz-Output_ParametersElements"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_8_3_SHPElementMesh.OnClick implementation

        ' Use the OpenFileDialog Class to choose VT files.

        ' Create a new ShapefileWorkspaceFactory CoClass to create a new workspace
        Dim workspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactoryClass

        ' System.IO.Path.GetDirectoryName(shapefileLocation) returns the directory part of the string. Example: "C:\test\"
        Dim featureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = CType(workspaceFactory.OpenFromFile(folderName, 0), ESRI.ArcGIS.Geodatabase.IFeatureWorkspace) ' Explicit Cast


        Dim templates As ESRI.ArcGIS.Framework.ITemplates = m_application.Templates
        Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra

        Dim openFileDialog As System.Windows.Forms.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        openFileDialog.InitialDirectory = "c:\"
        openFileDialog.Filter = "ArcArAz Files | *.dat;*.out"      '"Text Files (*.out)|*.out | Data Files  (*.dat)|*.dat"
        openFileDialog.FilterIndex = 2
        openFileDialog.Multiselect = True
        openFileDialog.RestoreDirectory = False
        openFileDialog.Title = "Select files generated by Visual ArcArAz"

        If openFileDialog.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            Dim filesVT() As String = openFileDialog.FileNames()
            Dim intNodes As String = " " 'Trim(Mid(strLine2, 1, 5))
            Dim intElements As String = " "  'Trim(Mid(strLine2, 6, 5))

            Dim pMxDoc As IMxDocument = m_application.Document
            Dim pMap As IMap = pMxDoc.ActiveView
            Dim pInLayer As IFeatureLayer = Nothing

            For s As Integer = 0 To pMap.LayerCount - 1
                If pMap.Layer(s).Name = "NodesMesh" Then
                    pInLayer = pMap.Layer(s)
                End If
            Next

            '''''''
            '''''''Create a feature class to store elements as polygons
            '''''''

            Dim strPathGeneral As String = System.IO.Path.GetDirectoryName(filesVT(0).ToString) 'sin la última contrabarra
            Dim strName As String = "ElementsMesh" '& Now()

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
                .GeometryType_2 = esriGeometryType.esriGeometryPolygon
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
                .Name_2 = "Node1"
                .Type_2 = esriFieldType.esriFieldTypeInteger
            End With
            pFieldsEdit.AddField(pField)

            ' Add another miscellaneous text field
            pField = New Field
            pFieldEdit = pField
            With pFieldEdit
                .Length_2 = 10
                .Name_2 = "Node2"
                .Type_2 = esriFieldType.esriFieldTypeInteger
            End With
            pFieldsEdit.AddField(pField)

            ' Add another miscellaneous text field
            pField = New Field
            pFieldEdit = pField
            With pFieldEdit
                .Length_2 = 10
                .Name_2 = "Node3"
                .Type_2 = esriFieldType.esriFieldTypeInteger
            End With
            pFieldsEdit.AddField(pField)

            ' Add another miscellaneous text field
            pField = New Field
            pFieldEdit = pField
            With pFieldEdit
                .Length_2 = 10
                .Name_2 = "Element"
                .Type_2 = esriFieldType.esriFieldTypeInteger
            End With
            pFieldsEdit.AddField(pField)
            ' Create the shapefile
            ' (some parameters apply to geodatabase options and can be defaulted as Nothing)
            Dim pPolyFeatClass As IFeatureClass = pFWS.CreateFeatureClass(strName, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, strShapeFieldName, "")

            'Add shapefile to TOC
            'Create a new ShapefileWorkspaceFactory object and open a shapefile folder

            Dim pFeatureWorkspace As IFeatureWorkspace = pWorkspaceFactory.OpenFromFile(strPathGeneral, 0)

            'Create a new FeatureLayer and assign a shapefile to it
            Dim pFeatureLayer As IFeatureLayer = New FeatureLayer
            pFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(strName)
            pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName

            'Add the FeatureLayer to the focus map
            pMap.AddLayer(pFeatureLayer)
            pMap.MoveLayer(pFeatureLayer, (pMxDoc.FocusMap.LayerCount))
            '''''''''''''''''
            '''''''''''''''''
            '''''''''''''''''

            Dim objGRIreader As StreamReader = Nothing
            For Each f As String In filesVT
                Dim strNameDAT As String = System.IO.Path.GetFileNameWithoutExtension(f)
                If strNameDAT.Contains("gri") Then '= True Or strNameDAT.Contains("gri1") = True Then
                    objGRIreader = New StreamReader(f)
                End If
            Next
            Dim strLineElements As String = objGRIreader.ReadLine()
            intElements = Trim(Mid(strLineElements, 10, 6))
            Dim strLineNodes As String = objGRIreader.ReadLine()
            intNodes = Trim(Mid(strLineNodes, 10, 6))
            Dim strEmptyLine As String = objGRIreader.ReadLine()

            Dim psbar As IStatusBar = m_application.StatusBar
            Dim pPro As IStepProgressor = psbar.ProgressBar

            Dim intNumLine As Integer = 3
            pPro.MinRange = 3 + 2 * intNodes
            pPro.MaxRange = 3 + 2 * intNodes + 1 + intElements
            pPro.StepValue = intElements / 100
            pPro.Step()
            pPro.Show()

            Dim pFeatClass As IFeatureClass = pInLayer.FeatureClass
            Dim pSearchFeatureBuffer As IFeatureBuffer = pFeatClass.CreateFeatureBuffer
            Dim pSearchFeatureCursor As IFeatureCursor = Nothing

            Dim pInsertFeatureBuffer As IFeatureBuffer = pPolyFeatClass.CreateFeatureBuffer
            Dim pInsertFeatureCursor As IFeatureCursor = pPolyFeatClass.Insert(True)
            Dim NewFeatureCount As Integer = 0

            Do Until intNumLine = 3 + 2 * intNodes + 1 + intElements
                Dim strGRIline As String = objGRIreader.ReadLine()
                If intNumLine > (3 + 2 * intNodes) Then


                    '' Define a constraint on the features to be deleted.
                    'Dim queryFilter As IQueryFilter = New QueryFilterClass With _
                    '                                  {.WhereClause = "TERM = 'Village'"}

                    '' Create a ComReleaser for cursor management.
                    'Using comReleaser As ComReleaser = New ComReleaser()
                    '    ' Create and manage a cursor.
                    '    Dim searchCursor As IFeatureCursor = FeatureClass.Search(queryFilter, False)
                    '    comReleaser.ManageLifetime(searchCursor)

                    '    ' Delete the retrieved features.
                    '    Dim feature As IFeature = searchCursor.NextFeature()
                    '    While Not feature Is Nothing
                    '        feature.Delete()
                    '        feature = searchCursor.NextFeature()
                    '    End While
                    'End Using


                    Dim i As Integer = CInt(Trim(Mid(strGRIline, 16, 5)))
                    Dim j As Integer = CInt(Trim(Mid(strGRIline, 21, 5)))
                    Dim k As Integer = CInt(Trim(Mid(strGRIline, 26, 5)))

                    Dim queryFilter As IQueryFilter = New QueryFilterClass With _
                                  {.WhereClause = "Node = " & i & " or Node = " & j & " or Node = " & k}

                    pSearchFeatureCursor = pFeatClass.Search(queryFilter, True)

                    'Build a polygon from a sequence of points.
                    'Add arrays of points to a geometry using the IGeometryBridge2 interface on the
                    'GeometryEnvironment singleton object.
                    Dim geometryBridge2 As IGeometryBridge2 = New GeometryEnvironmentClass
                    Dim pointCollection4 As IPointCollection4 = New PolygonClass

                    'TODO:
                    'pointCollection4.SpatialReference = 'Define the spatial reference of the new polygon.

                    Dim aWKSPointBuffer(2) As WKSPointZ
                    'Dim cPoints As Long = 3 'The number of points in the first part.
                    'ReDim aWKSPointBuffer(0 To CInt(cPoints - 1))
                    pSearchFeatureBuffer = pSearchFeatureCursor.NextFeature()
                    Dim pPoint1 As IPoint = pSearchFeatureBuffer.Shape
                    aWKSPointBuffer(0).X = pPoint1.X
                    aWKSPointBuffer(0).Y = pPoint1.Y
                    aWKSPointBuffer(0).Z = pSearchFeatureBuffer.Value(3)

                    pSearchFeatureBuffer = pSearchFeatureCursor.NextFeature()
                    Dim pPoint2 As IPoint = pSearchFeatureBuffer.Shape
                    aWKSPointBuffer(1).X = pPoint2.X
                    aWKSPointBuffer(1).Y = pPoint2.Y
                    aWKSPointBuffer(1).Z = pSearchFeatureBuffer.Value(3)

                    pSearchFeatureBuffer = pSearchFeatureCursor.NextFeature()
                    Dim pPoint3 As IPoint = pSearchFeatureBuffer.Shape
                    aWKSPointBuffer(2).X = pPoint3.X
                    aWKSPointBuffer(2).Y = pPoint3.Y
                    aWKSPointBuffer(2).Z = pSearchFeatureBuffer.Value(3)
                    'TODO:
                    'aWKSPointBuffer = 'Read cPoints into the point buffer.


                    'Dim queryFilteri As IQueryFilter = New QueryFilterClass With _
                    '              {.WhereClause = "Node = " & i & " and Node = " & j & " and Node = " & j}
                    'pSearchFeatureCursor = pFeatClass.Search(queryFilteri, True)
                    'Dim pFeat1 As IFeature = pSearchFeatureCursor.NextFeature()

                    'Dim queryFilterj As IQueryFilter = New QueryFilterClass With _
                    '              {.WhereClause = "Node = " & j}
                    'pSearchFeatureCursor = pFeatClass.Search(queryFilterj, True)
                    'Dim pFeat2 As IFeature = pSearchFeatureCursor.NextFeature()

                    'Dim queryFilterk As IQueryFilter = New QueryFilterClass With _
                    '              {.WhereClause = "Node = " & k}
                    'pSearchFeatureCursor = pFeatClass.Search(queryFilterk, True)
                    'Dim pFeat3 As IFeature = pSearchFeatureCursor.NextFeature()
                    'Dim pFeat1 As IFeature = pFeatClass.GetFeature(i - 1)
                    'Dim pFeat2 As IFeature = pFeatClass.GetFeature(j - 1)
                    'Dim pFEat3 As IFeature = pFeatClass.GetFeature(k - 1)

                    'Dim pPoint1 As IPoint = pFeat1.Shape
                    'Dim pPoint2 As IPoint = pFeat2.Shape
                    'Dim pPoint3 As IPoint = pFeat3.Shape

                    ''Build a polygon from a sequence of points.
                    ''Add arrays of points to a geometry using the IGeometryBridge2 interface on the
                    ''GeometryEnvironment singleton object.
                    'Dim geometryBridge2 As IGeometryBridge2 = New GeometryEnvironmentClass
                    'Dim pointCollection4 As IPointCollection4 = New PolygonClass

                    ''TODO:
                    ''pointCollection4.SpatialReference = 'Define the spatial reference of the new polygon.

                    'Dim aWKSPointBuffer(2) As WKSPoint
                    ''Dim cPoints As Long = 3 'The number of points in the first part.
                    ''ReDim aWKSPointBuffer(0 To CInt(cPoints - 1))
                    'aWKSPointBuffer(0).X = pPoint1.X
                    'aWKSPointBuffer(0).Y = pPoint1.Y
                    'aWKSPointBuffer(1).X = pPoint2.X
                    'aWKSPointBuffer(1).Y = pPoint2.Y
                    'aWKSPointBuffer(2).X = pPoint3.X
                    'aWKSPointBuffer(2).Y = pPoint3.Y
                    ''TODO:
                    ''aWKSPointBuffer = 'Read cPoints into the point buffer.
                    Dim pZAwarePoly As IZAware = pointCollection4
                    pZAwarePoly.ZAware = True

                    geometryBridge2.SetWKSPointZs(pointCollection4, aWKSPointBuffer)
                    Dim pPolygon As IPolygon = pointCollection4
                    pPolygon.Close()
                    'pointCollection4 has now been defined.


                    'Dim pLine1 As ILine = New Line
                    'pLine1.PutCoords(pPoint1, pPoint2)
                    'Dim pLine2 As ILine = New Line
                    'pLine2.PutCoords(pPoint2, pPoint3)
                    'Dim pLine3 As ILine = New Line
                    'pLine3.PutCoords(pPoint3, pPoint1)

                    'Dim pSegPoly As IPolygon = New Polygon
                    'Dim pZAwarePoly As IZAware = pSegPoly
                    'pZAwarePoly.ZAware = True
                    'Dim pRing As ISegmentCollection = New Ring
                    'pRing.AddSegment(pLine1)
                    'pRing.AddSegment(pLine2)
                    'pRing.AddSegment(pLine3)

                    'Dim pGeomColl As IGeometryCollection = pSegPoly
                    'pGeomColl.AddGeometry(pRing)

                    pInsertFeatureBuffer.Value(2) = i
                    pInsertFeatureBuffer.Value(3) = j
                    pInsertFeatureBuffer.Value(4) = k
                    pInsertFeatureBuffer.Value(5) = CInt(Trim(Mid(strGRIline, 1, 5)))

                    pInsertFeatureBuffer.Shape = pPolygon 'pFeature.Shape
                    pInsertFeatureCursor.InsertFeature(pInsertFeatureBuffer)

                    'Dim pFeat As IFeature = pPolyFeatClass.CreateFeature
                    'pFeat.Shape = pGeomColl
                    'pFeat.Value(2) = i
                    'pFeat.Value(3) = j
                    'pFeat.Value(4) = k
                    'pFeat.Value(5) = CInt(Trim(Mid(strGRIline, 1, 5)))
                    'pFeat.Store()
                    pSearchFeatureCursor.Flush()
                End If
                intNumLine = intNumLine + 1
                pPro.Position = intNumLine
                NewFeatureCount = NewFeatureCount + 1
                'Flush the feature cursor every 100 features
                'This is safer because you can write code to handle a flush error
                'If you don't flush the feature cursor it will automatically flush but
                'after all of your code executes at which time you have no control
                If NewFeatureCount = 100 Then
                    pInsertFeatureCursor.Flush()
                    NewFeatureCount = 0
                End If
            Loop
            pInsertFeatureCursor.Flush()
            pPro.Hide()
            pMxDoc.ActiveView.ContentsChanged()
            pMxDoc.UpdateContents()
        End If

    End Sub

End Class



