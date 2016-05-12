Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports System.Windows.Forms
Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem


<ComClass(cmd1SetWorkFolder.ClassId, cmd1SetWorkFolder.InterfaceId, cmd1SetWorkFolder.EventsId), _
 ProgId("ArcArAz.cmd1SetWorkFolder")> _
Public NotInheritable Class cmd1SetWorkFolder
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "54ce5f8c-3bf5-4dc1-bfa8-3c5bfce0df0e"
    Public Const InterfaceId As String = "3412b7d2-04f0-4ea1-ab2e-305b5a6e602c"
    Public Const EventsId As String = "5f12f788-062c-49e7-b9ee-b50f1468f817"
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
        MyBase.m_caption = "Set Project and WorkFolder"   'localizable text 
        MyBase.m_message = "Create a new folder to store your project"   'localizable text 
        MyBase.m_toolTip = "Create a new folder to store your project" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_SetWorkFolderCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd1SetWorkFolder.OnClick implementation
        'Dim folderBrowserDialog1 As FolderBrowserDialog = New FolderBrowserDialog
        'folderBrowserDialog1.Description = "Create of select a folder to store your data:"
        'folderBrowserDialog1.RootFolder = Environment.SpecialFolder.Desktop
        'Dim result As DialogResult = folderBrowserDialog1.ShowDialog()

        'If (result = DialogResult.OK) Then
        '    folderName = folderBrowserDialog1.SelectedPath
        'End If

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        pMxDoc.RelativePaths = True

        Dim commandBars As ICommandBars = m_application.Document.CommandBars
        Dim uid As UID = New UID
        uid.Value = "esriArcMapUI.MxFileMenuItem"
        uid.SubType = 4
        Dim commanditem As ICommandItem = commandBars.Find(uid, False, False)
        commanditem.Execute()


        'Get the path of the .mxd document
        Dim templates As ESRI.ArcGIS.Framework.ITemplates = m_application.Templates
        Dim pathMxd As String = templates.Item(templates.Count - 1) 'devuelve el string desde C:\ hasta .mxd
        folderName = System.IO.Path.GetDirectoryName(pathMxd) 'devuelve el string desde C:\ hasta el nombre de la carpeta sin la última contrabarra



    End Sub
    Public Sub FindCommandAndExecute(ByVal application As ESRI.ArcGIS.Framework.IApplication, ByVal commandName As System.String)

        Dim commandBars As ESRI.ArcGIS.Framework.ICommandBars = application.Document.CommandBars
        Dim uid As ESRI.ArcGIS.esriSystem.UID = New ESRI.ArcGIS.esriSystem.UIDClass
        uid.Value = commandName ' Example: "esriFramework.HelpContentsCommand" or "{D74B2F25-AC90-11D2-87F8-0000F8751720}"
        Dim commandItem As ESRI.ArcGIS.Framework.ICommandItem = commandBars.Find(uid, False, False)
        If Not (commandItem Is Nothing) Then
            commandItem.Execute()
        End If

    End Sub

    Public Function GetCommandOnToolbar(ByVal application As ESRI.ArcGIS.Framework.IApplication, ByVal toolbarName As System.String, ByVal commandName As System.String) As ESRI.ArcGIS.Framework.ICommandItem

        Dim commandBars As ESRI.ArcGIS.Framework.ICommandBars = application.Document.CommandBars
        Dim barID As ESRI.ArcGIS.esriSystem.UID = New ESRI.ArcGIS.esriSystem.UIDClass
        barID.Value = toolbarName ' Example: "esriArcMapUI.StandardToolBar"
        Dim barItem As ESRI.ArcGIS.Framework.ICommandItem = commandBars.Find(barID, False, False)

        If Not (barItem Is Nothing) AndAlso barItem.Type = ESRI.ArcGIS.Framework.esriCommandTypes.esriCmdTypeToolbar Then

            Dim commandBar As ESRI.ArcGIS.Framework.ICommandBar = CType(barItem, ESRI.ArcGIS.Framework.ICommandBar)
            Dim commandID As ESRI.ArcGIS.esriSystem.UID = New ESRI.ArcGIS.esriSystem.UIDClass
            commandID.Value = commandName ' Example: "esriArcMapUI.AddDataCommand"
            Return commandBar.Find(commandID, False)

        Else

            Return Nothing

        End If

    End Function
End Class



