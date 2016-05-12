Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.ADF

<ComClass(cmd1_2SetCoordinateSystem.ClassId, cmd1_2SetCoordinateSystem.InterfaceId, cmd1_2SetCoordinateSystem.EventsId), _
 ProgId("ArcArAz.cmd1_2SetCoordinateSystem")> _
Public NotInheritable Class cmd1_2SetCoordinateSystem
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "02e86e67-159b-40c3-9de9-7f92e4ab1e2f"
    Public Const InterfaceId As String = "c924f075-3acc-462c-ae73-6ba6fd0edc39"
    Public Const EventsId As String = "640f492d-f2c9-467c-a4b0-b94a9dd94b6a"
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
        MyBase.m_caption = "Set Coordinate System"   'localizable text 
        MyBase.m_message = "Define the coordinate system to work with"   'localizable text 
        MyBase.m_toolTip = "Define the coordinate system to work with" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_SetCoordinateSystemCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        ''TODO: Add cmd1_2SetCoordinateSystem.OnClick implementation
        'Dim spatialReferenceDialog As ISpatialReferenceDialog2 = New SpatialReferenceDialogClass
        'Dim hWnd As System.Int32
        'Dim spatialReference As ISpatialReference = spatialReferenceDialog.DoModalCreate(True, False, False, hWnd)

        'Dim pMxDoc As IMxDocument = m_application.Document
        'Dim pMap As IMap = pMxDoc.FocusMap

        'If (Not (pMap.SpatialReferenceLocked)) Then

        '    pMap.SpatialReference = spatialReference

        'End If

        'Check to see if ArcMap is in pagelayout view, else set it

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap

        'Step 1.
        Dim myPropertySheet As IComPropertySheet = New ComPropertySheetClass()
        myPropertySheet.Title = "My Property Sheet"
        myPropertySheet.HideHelpButton = True
        myPropertySheet.ActivePage = 1
        myPropertySheet.Title = "Coordinate System for " & pMap.Name

        'Step 2 and Step 3: Pass in layer, active view, and the application.
        Dim propertyObjects As ISet = New SetClass()
        propertyObjects.Add(pMxDoc.FocusMap)

        ''Step 4 - Bullet item b: Add page by page.
        myPropertySheet.ClearCategoryIDs()
        myPropertySheet.AddCategoryID(New UIDClass()) 'A dummy empty UID.
        myPropertySheet.AddPage(New ESRI.ArcGIS.CartoUI.MapProjectionPropPage) 'Feature layer symbology.

        'Step 5 - Show the property sheet.
        If myPropertySheet.CanEdit(propertyObjects) Then
            myPropertySheet.EditProperties(propertyObjects, m_application.hWnd)
        End If

    End Sub
End Class



