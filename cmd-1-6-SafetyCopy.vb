Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Geodatabase


<ComClass(cmd_1_6_SafetyCopy.ClassId, cmd_1_6_SafetyCopy.InterfaceId, cmd_1_6_SafetyCopy.EventsId), _
 ProgId("ArcArAz.cmd_1_6_SafetyCopy")> _
Public NotInheritable Class cmd_1_6_SafetyCopy
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "ab2748cd-ffa1-4d68-aa01-09fd165b3883"
    Public Const InterfaceId As String = "a0370fb3-cf97-4556-934d-66ea0e0400d6"
    Public Const EventsId As String = "f1dce032-a0b1-486a-9de9-680fbbd5483c"
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
        MyBase.m_caption = "Safety copy of shapefiles"   'localizable text 
        MyBase.m_message = "Create a safety copy of selected shapefiles in TOC"   'localizable text 
        MyBase.m_toolTip = "Create a safety copy of selected shapefiles in TOC" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_SafetyCopyCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_1_6_SafetyCopy.OnClick implementation
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

        ' Create a new ShapefileWorkspaceFactory CoClass to create a new workspace


        'Dim templates As ITemplates = m_application.Templates
        'Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        'folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra


        Dim intMes As Integer = Date.Now.Month
        Dim strMes As String = Nothing
        Select Case intMes
            Case 1
                strMes = "Jan"
            Case 2
                strMes = "Feb"
            Case 3
                strMes = "Mar"
            Case 4
                strMes = "Apr"
            Case 5
                strMes = "May"
            Case 6
                strMes = "Jun"
            Case 7
                strMes = "Jul"
            Case 8
                strMes = "Aug"
            Case 9
                strMes = "Sep"
            Case 10
                strMes = "Oct"
            Case 11
                strMes = "Nov"
            Case 12
                strMes = "Dic"
        End Select
        Dim fecha As String = strMes & Date.Now.Day.ToString & "h" & Date.Now.Hour.ToString & "m" & Date.Now.Minute.ToString

        Dim nameSecFolder As String
        For i = 0 To pLayers.Count - 1
            pLayer = pLayers.Element(i)
            Dim pFeatureLayer As IFeatureLayer = pLayer
            Dim featureClass As IFeatureClass = pFeatureLayer.FeatureClass
            Dim pDataset As IDataset = featureClass
            Dim pWorkspace As IWorkspace = pDataset.Workspace
            folderName = pWorkspace.PathName
            nameSecFolder = folderName & "\SafetyCopy"
            If System.IO.Directory.Exists(nameSecFolder) = False Then
                System.IO.Directory.CreateDirectory(nameSecFolder)
            End If


            Dim workspaceFactory As ESRI.ArcGIS.Geodatabase.IWorkspaceFactory = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactoryClass
            Dim featureWorkspace As IFeatureWorkspace = CType(workspaceFactory.OpenFromFile(nameSecFolder, 0), ESRI.ArcGIS.Geodatabase.IFeatureWorkspace) ' Explicit Cast

            Dim shpname As String = featureClass.AliasName & ".*"
            Dim filesofshape() As String = System.IO.Directory.GetFiles(folderName, shpname)

            For Each f As String In filesofshape
                If System.IO.File.Exists(nameSecFolder & "\" & System.IO.Path.GetFileNameWithoutExtension(f.ToString) & "_" & fecha & Right(f.ToString, 4)) = True Then
                    System.IO.File.Delete(nameSecFolder & "\" & System.IO.Path.GetFileNameWithoutExtension(f.ToString) & "_" & fecha & Right(f.ToString, 4))
                End If
                System.IO.File.Copy(f, nameSecFolder & "\" & System.IO.Path.GetFileNameWithoutExtension(f.ToString) & "_" & fecha & Right(f.ToString, 4))
            Next
        Next
    End Sub
End Class



