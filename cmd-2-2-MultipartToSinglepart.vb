Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.SystemUI

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Text
Imports ESRI.ArcGIS.Geoprocessor
Imports ESRI.ArcGIS.AnalysisTools
Imports ESRI.ArcGIS.DataManagementTools
Imports ESRI.ArcGIS.GeoprocessingUI



<ComClass(cmd2_2MultipartToSinglepart.ClassId, cmd2_2MultipartToSinglepart.InterfaceId, cmd2_2MultipartToSinglepart.EventsId), _
 ProgId("ArcArAz.cmd2_2MultipartToSinglepart")> _
Public NotInheritable Class cmd2_2MultipartToSinglepart
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4a15c147-1c8b-455b-9fd6-267cf824c138"
    Public Const InterfaceId As String = "abc77afa-fd1e-4c08-b0e7-cbfaf48740c9"
    Public Const EventsId As String = "030d6776-ff4c-4326-958e-ef9cfa5c6276"
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
        MyBase.m_category = "ArcArAz-EntityProp"  'localizable text 
        MyBase.m_caption = "Multipart to Singlepart"   'localizable text 
        MyBase.m_message = "Explode aggregate polygons into individual items"   'localizable text 
        MyBase.m_toolTip = "Explode aggregate polygons into individual items" 'localizable text 
        MyBase.m_name = "ArcArAz-EntityProp_MultipartToSinglepartCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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

        'Set a reference to the IGPCommandHelper2 interface.
        Dim pToolHelper As IGPToolCommandHelper2 = New GPToolCommandHelper

        'Set the tool you want to invoke.
        Dim toolboxPath = "C:\Program Files (x86)\ArcGIS\Desktop10.0\ArcToolbox\Toolboxes\Data Management Tools.tbx"
        Try
            pToolHelper.SetToolByName(toolboxPath, "MultipartToSinglepart")

            'Create the messages object to pass to the InvokeModal method.
            Dim msgs As IGPMessages
            msgs = New GPMessages

            'Invoke the tool.
            pToolHelper.InvokeModal(0, Nothing, True, msgs)
            m_application.CurrentTool = Nothing
        Catch err As Exception
            MsgBox("The tool instalation folder is not the default." & vbCrLf & _
                    "Please, access this tool throught ArcToolBox: " & vbCrLf & _
                    "ArcToolBox / Data Management Tools / Features / Multipart To Singlepart.")
        End Try



        '        ' ''Dim pMxDoc As IMxDocument = m_application.Document
        '        ' ''Dim pMap As IMap = pMxDoc.FocusMap
        '        ' ''Dim pInLayer As IFeatureLayer = pMxDoc.SelectedLayer

        '        ' ''If pInLayer Is Nothing Then  'Check if no input layer is selected
        '        ' ''    MsgBox("Select a POLYGON feature class in the TOC to check existing holes and remove them", vbCritical, "Incompatible input layer")
        '        ' ''    Exit Sub
        '        ' ''End If

        '        ' ''Dim pInFeatClass As IFeatureClass = pInLayer.FeatureClass

        '        ' ''If Not pInFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
        '        ' ''    MsgBox("Select a POLYGON feature class in the TOC to check existing holes and remove them", vbCritical, "Incompatible input layer")
        '        ' ''    Exit Sub
        '        ' ''End If

        '        ' ''For i As Integer = 0 To pInFeatClass.FeatureCount(Nothing) - 1
        '        ' ''    Dim pfeat As IFeature = pInFeatClass.GetFeature(i)
        '        ' ''    Dim pGeom As IGeometryCollection = pfeat.Shape
        '        ' ''    If pGeom.GeometryCount > 1 Then
        '        ' ''        pGeom.
        '        ' ''    End If
        '        ' ''Next

        '        Dim pClone As IClone
        '        Dim pDataset As IDataset
        '        Dim pFeature As IFeature
        '        Dim pFeatureClass As IFeatureClass
        '        Dim pFeatureCursor As IFeatureCursor
        '        Dim pFeatureLayer As IFeatureLayer
        '        Dim pFeatureWorkspace As IFeatureWorkspace
        '        Dim pFields As IFields
        '        Dim pGeometryColl As IGeometryCollection
        '        Dim pInsertFeatureBuffer As IFeatureBuffer
        '        Dim pInsertFeatureCursor As IFeatureCursor
        '        Dim pNewFeatureClass As IFeatureClass
        '        Dim pPolygon As IPolygon4
        '        Dim pPolygonArray() As IPolygon

        '        Dim strNewFeatureClassName As String
        '        Dim GeometryCount As Integer
        '        Dim lShapeFieldIndex As Long

        '        On Error GoTo ErrorHandler

        '        Dim pMxDoc As IMxDocument = m_application.Document
        '        'Make certain the selected item in the toc is a feature layer
        '        If pMxDoc.SelectedItem Is Nothing Then
        '            MsgBox("Select a feature layer in the table of contents as the input feature class.")
        '            Exit Sub
        '        End If

        '        If Not TypeOf pMxDoc.SelectedItem Is IFeatureLayer Then
        '            MsgBox("No feature layer selected.")
        '            Exit Sub
        '        End If

        '        pFeatureLayer = pMxDoc.SelectedItem
        '        pFeatureClass = pFeatureLayer.FeatureClass

        '        'Don't process point layers, they have no multi-part features
        '        If pFeatureClass.ShapeType = esriGeometryType.esriGeometryPoint Then
        '            MsgBox("Point layers do not have multi-parts.  Exiting.")
        '            Exit Sub
        '        End If


        '        strNewFeatureClassName = InputBox("Enter New Shapefile name:", "New Shapefile")
        '        If strNewFeatureClassName = "" Then Exit Sub

        '        'Create a new feature class to store the new features
        '        'Create the feature class in the same dataset if one exists - shapefiles don't have one
        '        pFields = pFeatureLayer.FeatureClass.Fields
        '        If pFeatureClass.FeatureDataset Is Nothing Then
        '            pDataset = pFeatureClass
        '            pFeatureWorkspace = pDataset.Workspace
        '            pNewFeatureClass = pFeatureWorkspace.CreateFeatureClass(strNewFeatureClassName, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
        '        Else
        '            pNewFeatureClass = pFeatureClass.FeatureDataset.CreateFeatureClass(strNewFeatureClassName, pFields, Nothing, Nothing, esriFeatureType.esriFTSimple, pFeatureClass.ShapeFieldName, "")
        '        End If

        '        'Create an insert cursor
        '        pInsertFeatureCursor = pNewFeatureClass.Insert(True)
        '        pInsertFeatureBuffer = pNewFeatureClass.CreateFeatureBuffer

        '        'Copy each feature from the original feature class to the new feature class
        '        pFeatureCursor = pFeatureClass.Search(Nothing, True)
        '        pFeature = pFeatureCursor.NextFeature
        '        Do While Not pFeature Is Nothing
        '            pGeometryColl = pFeature.Shape
        '            If pGeometryColl.GeometryCount = 1 Then
        '                'Single part feature, straight copy
        '                InsertFeature(pInsertFeatureCursor, pInsertFeatureBuffer, pFeature, pFeature.Shape)
        '            ElseIf pFeature.Shape.GeometryType = esriGeometryType.esriGeometryPolygon Then
        '                pPolygon = pFeature.Shape
        '                Dim pGeomColl As IGeometryCollection = pPolygon

        '                'Dim lngRingCount As Long = pPolygon.ExteriorRingCount
        '                'Dim pPolyAr As IPolygon = pPolygonArray(0)
        '                'pPolygon.GetConnectedComponents(pPolygon.ExteriorRingCount, pPolygonArray())

        '                For GeometryCount = 0 To pPolygon.ExteriorRingCount - 1
        '                    InsertFeature(pInsertFeatureCursor, pInsertFeatureBuffer, pFeature, pGeomColl.Geometry(GeometryCount)) 'pPolygonArray(GeometryCount))
        '                Next GeometryCount
        '            Else
        '                'Multipart feature, create a new feature from each part
        '                For GeometryCount = 0 To pGeometryColl.GeometryCount - 1
        '                    InsertFeature(pInsertFeatureCursor, pInsertFeatureBuffer, pFeature, pGeometryColl.Geometry(GeometryCount))
        '                Next GeometryCount
        '            End If

        '            'Get the next feature
        '            pFeature = pFeatureCursor.NextFeature
        '        Loop

        '        Exit Sub 'Exit sub to avoid error handler

        'ErrorHandler:
        '        MsgBox("An error occurred. Check that the shapefile specified doesn't already exist.")
        '        Exit Sub
    End Sub

    '    Private Sub InsertFeature(ByVal pInsertFeatureCursor As IFeatureCursor, ByVal pInsertFeatureBuffer As IFeatureBuffer, ByVal pOrigFeature As IFeature, ByVal pGeometry As IGeometry)
    '        Dim pGeometryColl As IGeometryCollection
    '        Dim pFields As IFields
    '        Dim pField As IField
    '        Dim pPointColl As IPointCollection
    '        Dim FieldCount As Integer

    '        'Copy the attributes of the orig feature the new feature
    '        pFields = pOrigFeature.Fields
    '        For FieldCount = 0 To pFields.FieldCount - 1  'skip OID and geometry
    '            pField = pFields.Field(FieldCount)
    '            If Not pField.Type = esriFieldType.esriFieldTypeGeometry And Not pField.Type = esriFieldType.esriFieldTypeOID _
    '              And pField.Editable Then
    '                pInsertFeatureBuffer.Value(FieldCount) = pOrigFeature.Value(FieldCount)
    '            End If
    '        Next FieldCount

    '        'Handle cases where parts are passed down
    '        If pGeometry.GeometryType = esriGeometryType.esriGeometryPath Then
    '            pGeometryColl = New Polyline
    '            pGeometryColl.AddGeometries(1, pGeometry)
    '            pGeometry = pGeometryColl
    '        ElseIf pOrigFeature.Shape.GeometryType = esriGeometryType.esriGeometryMultipoint Then
    '            If TypeOf pGeometry Is IMultipoint Then
    '                pPointColl = pGeometry
    '                pGeometry = pPointColl.Point(0)
    '            End If
    '            pGeometryColl = New Multipoint
    '            pGeometryColl.AddGeometries(1, pGeometry)
    '            pGeometry = pGeometryColl
    '        End If

    '        pInsertFeatureBuffer.Shape = pGeometry
    '        pInsertFeatureCursor.InsertFeature(pInsertFeatureBuffer)
    '        pInsertFeatureCursor.Flush()

    '    End Sub


End Class



