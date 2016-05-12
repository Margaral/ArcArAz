Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase

<ComClass(cmd_1_7_RemoveShapefile.ClassId, cmd_1_7_RemoveShapefile.InterfaceId, cmd_1_7_RemoveShapefile.EventsId), _
 ProgId("ArcArAz.cmd_1_7_RemoveShapefile")> _
Public NotInheritable Class cmd_1_7_RemoveShapefile
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "c6d4a1b8-ef1b-49eb-9c2d-691389528e66"
    Public Const InterfaceId As String = "a863f1bb-a548-4d3e-aba0-35d5a9e5ae78"
    Public Const EventsId As String = "3e7eff27-f65f-4d74-ab57-48b4234dba02"
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
        MyBase.m_caption = "Remove shapefiles"   'localizable text 
        MyBase.m_message = "Remove selected shapefiles from TOC and WorkFolder"   'localizable text 
        MyBase.m_toolTip = "Remove selected shapefiles from TOC and WorkFolder" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_RemoveShpCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_1_7_RemoveShapefile.OnClick implementation

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim ppSet As ISet = Nothing
        Dim pLayer As ILayer = Nothing
        Dim pLayers As IArray = New Array

        'los contentsview son las distintas vistas del TOC, de las 4 opciones que hay, de list by order drawing...

        Dim pMxTOC As IContentsView3 = pMxDoc.CurrentContentsView
        Dim p As Object = pMxTOC.SelectedItem

        If TypeOf p Is ISet Then
            ppSet = p
            ppSet.Reset()
            For i = 0 To ppSet.Count
                pLayer = ppSet.Next
                If Not pLayer Is Nothing Then
                    pLayers.Add(pLayer)
                End If
            Next
        ElseIf TypeOf p Is ILayer Then
            pLayer = p
            pLayers.Add(pLayer)
        End If

        If pLayers.Count = 0 Then  'Check if no input layer is selected
            MsgBox("Select feature classes in TOC to create a safety copy", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        For i = 0 To pLayers.Count - 1
            pLayer = pLayers.Element(i)
            Dim pFeatureLayer As IFeatureLayer = pLayer
            Dim featureClass As IFeatureClass = pFeatureLayer.FeatureClass
            Dim pDataset As IDataset = featureClass
            Dim pWorkspace As IWorkspace = pDataset.Workspace
            folderName = pWorkspace.PathName
            Dim workspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactoryClass
            Dim featureWorkspace As IFeatureWorkspace = CType(workspaceFactory.OpenFromFile(folderName, 0), ESRI.ArcGIS.Geodatabase.IFeatureWorkspace) ' Explicit Cast
            pMap.DeleteLayer(pFeatureLayer)

            Dim shpname As String = featureClass.AliasName & ".*"
            Dim filesofshape() As String = System.IO.Directory.GetFiles(folderName, shpname)

            For Each f As String In filesofshape
                System.IO.File.Delete(folderName & "\" & System.IO.Path.GetFileName(f.ToString))
                On Error Resume Next
            Next
        Next

        pMxDoc.ActiveView.ContentsChanged()
        pMxDoc.UpdateContents()
    End Sub
End Class



