Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports System.IO
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.GeoprocessingUI

<ComClass(cmd_3_4_RasterTINToASCII.ClassId, cmd_3_4_RasterTINToASCII.InterfaceId, cmd_3_4_RasterTINToASCII.EventsId), _
 ProgId("ArcArAz.cmd_3_4_RasterTINToASCII")> _
Public NotInheritable Class cmd_3_4_RasterTINToASCII
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "564ebb82-9e5f-4f2c-b840-bf8e2f0967bd"
    Public Const InterfaceId As String = "e3cffefc-0461-4376-9216-c5b23444ade3"
    Public Const EventsId As String = "c91bae4e-20a5-4c6a-afc6-8f564007d36b"
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
        MyBase.m_caption = "Raster/TIN to ASCII"   'localizable text 
        MyBase.m_message = "Create a .txt file with the X, Y and Z coordinates of raster cells or TIN nodes"   'localizable text 
        MyBase.m_toolTip = "Create a .txt file with the X, Y and Z coordinates of raster cells or TIN nodes" 'localizable text 
        MyBase.m_name = "ArcArAz-SpatialFields_RasterTINToASCIICommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_3_4_RasterTINToASCII.OnClick implementation
        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pInLayer As ILayer = pMxDoc.SelectedLayer

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a TIN or RASTER layer in the TOC", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        If TypeOf pInLayer Is ITinLayer Then  'check if selected layer is a feature layer
            'set selected layer as input feature layer
            Dim pTINLayer As ITinLayer = pMxDoc.SelectedLayer
            Dim pDataSet As IDataset = pTINLayer.Dataset
            Dim strTinPath As String = Left(pDataSet.BrowseName, InStrRev(pDataSet.BrowseName, "\"))
            Dim strTxtName = pInLayer.Name & "_TIN2ASCII.txt"
            Dim strTxtWholeName As String = strTinPath & strTxtName

            If File.Exists(strTxtWholeName) Then
                File.Delete(strTxtWholeName)
            End If

            Dim pTin As ITin = pTINLayer.Dataset
            Dim pNodeColl As ITinNodeCollection = pTin

            Dim psbar As IStatusBar
            psbar = m_application.StatusBar
            Dim pPro As IStepProgressor
            pPro = psbar.ProgressBar

            pPro.MinRange = 1
            pPro.MaxRange = pNodeColl.NodeCount
            pPro.StepValue = pNodeColl.NodeCount / 100
            pPro.Step()
            pPro.Show()

            Dim objWriter As StreamWriter = New StreamWriter(strTxtWholeName)

            For i = 5 To pNodeColl.NodeCount
                objWriter.WriteLine(Math.Round(pNodeColl.GetNode(i).X, 1) & " " & Math.Round(pNodeColl.GetNode(i).Y, 1) & " " & Math.Round(pNodeColl.GetNode(i).Z, 3))
                pPro.Position = i
            Next
            objWriter.Close()
            pPro.Hide()

        ElseIf TypeOf pInLayer Is IRasterLayer Then  'Raster2TXT
            Dim arcToolBoxExtension As IArcToolboxExtension = m_application.FindExtensionByName("ESRI ArcToolbox")

            If Not arcToolBoxExtension Is Nothing Then
                Dim arcToolBox As IArcToolbox = arcToolBoxExtension.ArcToolbox
                Dim gpTool = arcToolBox.GetToolbyNameString("Raster2TXT")
                If Not gpTool Is Nothing Then
                    arcToolBox.InvokeTool(m_application.hWnd, gpTool, Nothing, False)
                End If
            End If
            'MsgBox("Select a TIN layer in the TOC", vbCritical, "Incompatible input layer")

            'Dim pRasterExportOp As IRasterExporter
            'pRasterExportOp = New RasterConversionOp

            ''Get raster
            'Dim pRas01 As IRaster = readRasterFromDisk("myRaster")

            'Dim sOutFname As String = "myFile.txt"

            '' Call the export method
            'pRasterExportOp.ExportToASCII(pRas01, sOutFname)
            'Dim GP As IGeoProcessor = New GeoProcessor


            Exit Sub
        End If
        MsgBox("Process finished at " & Now, MsgBoxStyle.Information)
    End Sub
End Class



