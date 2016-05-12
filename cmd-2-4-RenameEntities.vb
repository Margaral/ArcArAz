Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto

<ComClass(cmd2_4RenameEntities.ClassId, cmd2_4RenameEntities.InterfaceId, cmd2_4RenameEntities.EventsId), _
 ProgId("ArcArAz.cmd2_4RenameEntities")> _
Public NotInheritable Class cmd2_4RenameEntities
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "88662805-e92c-44de-85bf-96ea683748fe"
    Public Const InterfaceId As String = "243d3f98-ae90-4a48-a456-82dce78215c1"
    Public Const EventsId As String = "8d0b93e7-1275-4cf0-b3a1-01f654edd8aa"
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
        MyBase.m_caption = "Rename Entities"   'localizable text 
        MyBase.m_message = "Create secuential names to avoid weird characters, duplicated names or too long names"   'localizable text 
        MyBase.m_toolTip = "Create secuential names to avoid weird characters, duplicated names or too long names" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_RenameEntitiesCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd2_4RenameEntities.OnClick implementation
        Dim pMxDoc As IMxDocument
        pMxDoc = m_application.Document

        Dim pMap As IMap
        pMap = pMxDoc.FocusMap

        Dim pInLayer As IFeatureLayer
        pInLayer = pMxDoc.SelectedLayer

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a feature class in the TOC to create new names", vbCritical, "Input layer not set")
            Exit Sub
        End If

        Dim SelectFieldName As New cmd_2_4_Frm_1_SelectFieldName
        SelectFieldName.Application = m_application
        SelectFieldName.ShowDialog()


    End Sub
End Class



