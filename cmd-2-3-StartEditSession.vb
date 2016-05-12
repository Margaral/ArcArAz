Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor

<ComClass(cmd2_3StartEditSession.ClassId, cmd2_3StartEditSession.InterfaceId, cmd2_3StartEditSession.EventsId), _
 ProgId("ArcArAz.cmd2_3StartEditSession")> _
Public NotInheritable Class cmd2_3StartEditSession
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "8bb287ed-872e-40fd-a5e3-242201ff5e3c"
    Public Const InterfaceId As String = "10ac0115-f4f4-4807-b354-6ce579b2c708"
    Public Const EventsId As String = "048f74fe-afa0-48ca-bda4-aa3743145e34"
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
        MyBase.m_caption = "Start Edit Session"   'localizable text 
        MyBase.m_message = "Start an edit session to create, delete or modify entities"   'localizable text 
        MyBase.m_toolTip = "Start an edit session to create, delete or modify entities" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_StarEditSessionCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd2_3StartEditSession.OnClick implementation
        Dim commandBars As ESRI.ArcGIS.Framework.ICommandBars = m_application.Document.CommandBars
        Dim commandID As ESRI.ArcGIS.esriSystem.UID = New ESRI.ArcGIS.esriSystem.UIDClass
        Dim commandName As String = "esriEditor.EditingToolbarNew" '"esriEditor.EditorToolBar"
        commandID.Value = commandName ' Example: "esriFramework.HelpContentsCommand"

        Dim pCommandItem As ICommandItem = commandBars.Find(commandID, False, False)
        Dim pKRCommandBar As ICommandBar = pCommandItem

        pKRCommandBar.Dock(esriDockFlags.esriDockFloat)


        Dim pUID As New UID
        Dim pCmdItem As ICommandItem
        ' Use the GUID of the Save command
        pUID.Value = "{59D2AFD0-9EA2-11D1-9165-0080C718DF97}"
        ' or you can use the ProgID
        ' pUID.Value = "esriArcMapUI.MxFileMenuItem"
        'pUID.SubType = 3
        pCmdItem = commandBars.Find(pUID)
        pCmdItem.Execute()

        'Dim uid As ESRI.ArcGIS.esriSystem.UID
        'uid = New UIDClass()
        'uid.Value = "esriEditor.Editor"
        'Dim editor As IEditor
        'editor = CType(m_application.FindExtensionByCLSID(uid), IEditor)
        'editor.StartEditing()
        ''Check to see if a workspace is already being edited.
        'If editor.EditState = esriEditState.esriStateNotEditing Then
        '    editor.StartEditing(workspaceToEdit)
        '    Return True
        'Else
        '    Return False
        'End If
    End Sub
End Class



