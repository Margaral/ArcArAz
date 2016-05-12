Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.SystemUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.Geodatabase

<ComClass(cmd1_2AddShapefile.ClassId, cmd1_2AddShapefile.InterfaceId, cmd1_2AddShapefile.EventsId), _
 ProgId("ArcArAz.cmd1_2AddShapefile")> _
Public NotInheritable Class cmd1_2AddShapefile
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "5f914bd7-f9f1-4753-a603-d2c52a71b0c7"
    Public Const InterfaceId As String = "0f647614-65cb-4f5d-835b-1d79740b9efa"
    Public Const EventsId As String = "76679566-53de-4ba6-b2de-e23debdaaeb3"
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
        MyBase.m_caption = "Add Shapefile to TOC and WorkFolder"   'localizable text 
        MyBase.m_message = "Create a copy of the original file in the workfolder and add this copy to TOC"   'localizable text 
        MyBase.m_toolTip = "Create a copy of the original file in the workfolder and add this copy to TOC" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_AddShapefileCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd1_2AddShapefile.OnClick implementation
        ' Use the OpenFileDialog Class to choose which shapefile to load.

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim activeView As IMap = pMxDoc.ActiveView

        ' Create a new ShapefileWorkspaceFactory CoClass to create a new workspace
        Dim workspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactoryClass

        ' System.IO.Path.GetDirectoryName(shapefileLocation) returns the directory part of the string. Example: "C:\test\"
        Dim featureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = CType(workspaceFactory.OpenFromFile(folderName, 0), ESRI.ArcGIS.Geodatabase.IFeatureWorkspace) ' Explicit Cast


        Dim templates As ESRI.ArcGIS.Framework.ITemplates = m_application.Templates
        Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra

        Dim openFileDialog As System.Windows.Forms.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        openFileDialog.InitialDirectory = "c:\"
        openFileDialog.Filter = "Shapefiles (*.shp)|*.shp"
        openFileDialog.FilterIndex = 2
        openFileDialog.RestoreDirectory = False
        openFileDialog.Multiselect = True
        openFileDialog.RestoreDirectory = False


        If openFileDialog.ShowDialog = System.Windows.Forms.DialogResult.OK Then

            ' The user chose a particular shapefile.

            ' The returned string will be the full path, filename and file-extension for the chosen shapefile. Example: "C:\test\cities.shp"
            Dim shapefileLocation() As String = openFileDialog.FileNames()

            For Each shploc As String In shapefileLocation

                'Si se añade un shapefile que está en la carpeta del proyecto desde la carpeta del proyecto
                If System.IO.Path.GetDirectoryName(shploc.ToString) = folderName Then
                    Dim featureClass1 As IFeatureClass = featureWorkspace.OpenFeatureClass(System.IO.Path.GetFileNameWithoutExtension(shploc))

                    Dim featureLayer1 As IFeatureLayer = New FeatureLayerClass
                    featureLayer1.FeatureClass = featureClass1
                    featureLayer1.Name = featureClass1.AliasName
                    featureLayer1.Visible = True
                    pMxDoc.FocusMap.AddLayer(featureLayer1)

                    ' Zoom the display to the full extent of all layers in the map

                    pMxDoc.ActiveView.ContentsChanged()
                    pMxDoc.UpdateContents()

                    'Si se añade un shapefile que está en la carpeta del proyecto desde otra carpeta
                ElseIf System.IO.File.Exists(folderName & "\" & System.IO.Path.GetFileNameWithoutExtension(shploc) & ".shp") Then
                    MsgBox("There is a shapefile with the same name in your project folder. " & vbCrLf & _
                    "ArcArAz will add the shapefile from your project folder.")
                    Dim featureClass1 As IFeatureClass = featureWorkspace.OpenFeatureClass(System.IO.Path.GetFileNameWithoutExtension(shploc))

                    Dim featureLayer1 As IFeatureLayer = New FeatureLayerClass
                    featureLayer1.FeatureClass = featureClass1
                    featureLayer1.Name = featureClass1.AliasName
                    featureLayer1.Visible = True
                    pMxDoc.FocusMap.AddLayer(featureLayer1)

                    ' Zoom the display to the full extent of all layers in the map

                    pMxDoc.ActiveView.ContentsChanged()
                    pMxDoc.UpdateContents()
                Else
                    Dim shpname As String = System.IO.Path.GetFileNameWithoutExtension(shploc) & ".*"
                    Dim filesofshape() As String = System.IO.Directory.GetFiles(System.IO.Path.GetDirectoryName(shploc), shpname)

                    For Each f As String In filesofshape
                        System.IO.File.Copy(f, folderName & "\" & System.IO.Path.GetFileName(f))
                    Next

                    ' System.IO.Path.GetFileNameWithoutExtension(shapefileLocation) returns the base filename (without extension). Example: "cities"
                    Dim featureClass As IFeatureClass = featureWorkspace.OpenFeatureClass(System.IO.Path.GetFileNameWithoutExtension(shploc))

                    Dim featureLayer As IFeatureLayer = New FeatureLayerClass
                    featureLayer.FeatureClass = featureClass
                    featureLayer.Name = featureClass.AliasName
                    featureLayer.Visible = True
                    pMxDoc.FocusMap.AddLayer(featureLayer)

                    ' Zoom the display to the full extent of all layers in the map

                    'pMxDoc.ActiveView.Extent = activeView.FullExtent
                    'pMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGeography, Nothing, Nothing)
                    pMxDoc.ActiveView.ContentsChanged()
                    pMxDoc.UpdateContents()
                End If

            Next
        End If




    End Sub
End Class



