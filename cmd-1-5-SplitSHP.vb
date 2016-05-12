Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.GeoDatabaseUI
Imports ESRI.ArcGIS.Geometry

<ComClass(cmd_1_5_SplitSHP.ClassId, cmd_1_5_SplitSHP.InterfaceId, cmd_1_5_SplitSHP.EventsId), _
 ProgId("ArcArAz.cmd_1_5_SplitSHP")> _
Public NotInheritable Class cmd_1_5_SplitSHP
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "a08d93a6-4b33-4ec0-8184-437bb0fd04e3"
    Public Const InterfaceId As String = "690cf8df-21a7-4480-9690-99eb18110ad5"
    Public Const EventsId As String = "04b07556-6544-41d5-a16f-63240873d17c"
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
        MyBase.m_category = "ArcArAz-Input"  'localizable text 
        MyBase.m_caption = "Split Shapefile"   'localizable text 
        MyBase.m_message = "Select a shapefile to split into several ones"   'localizable text 
        MyBase.m_toolTip = "Select a shapefile to split into several ones" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_SplitSHPCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_1_5_SplitSHP.OnClick implementation
        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pInLayer As IFeatureLayer = pMxDoc.SelectedLayer

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a feature class in the TOC to split", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim pInFeatClass As IFeatureClass = pInLayer.FeatureClass

        Dim pFeatSel As IFeatureSelection = pInLayer

        Dim pFeatClass As IFeatureClass = pInLayer.FeatureClass

        Dim pDataset As IDataset = pFeatClass

        Dim sPath As String = pDataset.Workspace.PathName & "\"

        Dim pINFeatureClassName As IFeatureClassName = pDataset.FullName

        Dim pInDsName As IDatasetName
        pInDsName = pINFeatureClassName

        Dim Tamaño As Integer
        Tamaño = InputBox("Introduce the maximum number of entities for the new shapefiles")
        MkDir(sPath & "\Split_" & pDataset.Name & Tamaño)

        For j = 0 To pFeatClass.FeatureCount(Nothing) Step Tamaño '- 1 '""""16 agosto 2012

            For i = j To j + Tamaño - 1
                Dim pFeat As IFeature
                On Error Resume Next
                pFeat = pFeatClass.GetFeature(i)
                pFeatSel.Add(pFeat)
            Next

            Dim pSelSet As ISelectionSet
            pSelSet = pFeatSel.SelectionSet

            'Define the output feature class name

            Dim pFeatureClassName As IFeatureClassName
            pFeatureClassName = New FeatureClassName

            Dim pOutDatasetName As IDatasetName
            pOutDatasetName = pFeatureClassName

            pOutDatasetName.Name = "Export_" & j

            Dim pWorkspaceName As IWorkspaceName
            pWorkspaceName = New WorkspaceName

            pWorkspaceName.PathName = sPath & "\Split_" & pDataset.Name & Tamaño

            pWorkspaceName.WorkspaceFactoryProgID = "esriCore.shapefileworkspacefactory.1"

            pOutDatasetName.WorkspaceName = pWorkspaceName

            pFeatureClassName.FeatureType = esriFeatureType.esriFTSimple

            pFeatureClassName.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryAny

            pFeatureClassName.ShapeFieldName = "Shape"

            Dim pExportOp As IExportOperation
            pExportOp = New ExportOperation

            pExportOp.ExportFeatureClass(pInDsName, Nothing, pSelSet, Nothing, pOutDatasetName, 0)

            pFeatSel.Clear()

        Next
    End Sub
End Class



