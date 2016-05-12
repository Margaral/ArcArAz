Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.GeoprocessingUI
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.esriSystem

<ComClass(cmd1_3ImportTopoAutoCAD.ClassId, cmd1_3ImportTopoAutoCAD.InterfaceId, cmd1_3ImportTopoAutoCAD.EventsId), _
 ProgId("ArcArAz.cmd1_3ImportTopoAutoCAD")> _
Public NotInheritable Class cmd1_3ImportTopoAutoCAD
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "31d8e4a5-aa44-4bb7-b137-6a81aafc1dcb"
    Public Const InterfaceId As String = "27d6f26c-6cbc-4a9a-811a-02fd0d847d3f"
    Public Const EventsId As String = "15bddbcd-817e-494c-93b5-fb6b47fd2ddb"
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
        MyBase.m_caption = "Import from AutoCAD"   'localizable text 
        MyBase.m_message = "Import entities from an CAD file, .dwg or .dxf, to Shapefile"   'localizable text 
        MyBase.m_toolTip = "Import entities from an CAD file, .dwg or .dxf, to Shapefile" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_ImportFromAutoCADCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd1_3ImportTopoAutoCAD.OnClick implementation
        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pInLayer As ILayer = pMxDoc.SelectedLayer

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Add to TOC a CAD file and select it to use this tool", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim templates As ESRI.ArcGIS.Framework.ITemplates = m_application.Templates
        Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra

        CreateAccessWorkspace(folderName & "\", pInLayer.Name & ".mdb")
        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactoryClass()



        'Set a reference to the IGPCommandHelper2 interface.
        Dim pToolHelper As IGPToolCommandHelper2 = New GPToolCommandHelper

        'Set the tool you want to invoke.
        Dim toolboxPath = "C:\Program Files (x86)\ArcGIS\Desktop10.0\ArcToolbox\Toolboxes\Conversion Tools.tbx"
        Try
            pToolHelper.SetToolByName(toolboxPath, "CadToGeodatabase")

            'Create the messages object to pass to the InvokeModal method.
            Dim msgs As IGPMessages
            msgs = New GPMessages

            'Invoke the tool.
            pToolHelper.InvokeModal(0, Nothing, True, msgs)
            m_application.CurrentTool = Nothing
        Catch err As Exception
            MsgBox("The tool instalation folder is not the default." & vbCrLf & _
                    "Please, access this tool throught ArcToolBox: " & vbCrLf & _
                    "ArcToolBox / Conversion Tools / To Geodatabase / CAD to Geodatabase")
        End Try
    End Sub
    Public Function CreateAccessWorkspace(ByVal Path As String, ByVal NameGDB As String) As IWorkspace
        'workspaceFactory.Create("C:\temp\", "Sample.mdb", Nothing, 0)

        ' Instantiate an Access workspace factory and create a new personal geodatabase.
        ' The Create method returns a workspace name object.

        ' Delete an existing GDB
        If System.IO.File.Exists(Path & NameGDB) Then
            System.IO.File.Delete(Path & NameGDB)
        End If

        Dim workspaceFactory As IWorkspaceFactory = New AccessWorkspaceFactoryClass()
        Dim workspaceName As IWorkspaceName = workspaceFactory.Create(Path, NameGDB, Nothing, 0)

        ' Cast the workspace name object to the IName interface and open the workspace.
        Dim Name As IName = CType(workspaceName, IName)
        Dim workspace As IWorkspace = CType(Name.Open(), IWorkspace)
        Return workspace
    End Function
End Class



