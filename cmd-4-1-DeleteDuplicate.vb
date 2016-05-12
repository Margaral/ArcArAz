Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

<ComClass(cmd_4_1_DeleteDuplicate.ClassId, cmd_4_1_DeleteDuplicate.InterfaceId, cmd_4_1_DeleteDuplicate.EventsId), _
 ProgId("ArcArAz.cmd_4_1_DeleteDuplicate")> _
Public NotInheritable Class cmd_4_1_DeleteDuplicate
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4a7d87ad-3fd8-4a41-9a14-95571ae99feb"
    Public Const InterfaceId As String = "2910840c-c743-4b61-bbaf-4e3665a583d7"
    Public Const EventsId As String = "671e6266-3a2f-4431-9c97-525e3480399b"
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
        MyBase.m_category = "ArcArAz-Rivers"  'localizable text 
        MyBase.m_caption = "Delete Duplicate Segments"   'localizable text 
        MyBase.m_message = "Delete duplicate sections in a river shapefile"   'localizable text 
        MyBase.m_toolTip = "Delete duplicate sections in a river shapefile" 'localizable text 
        MyBase.m_name = "ArcArAz-Rivers_DeleteDuplicateCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_4_1_DeleteDuplicate.OnClick implementation
        Dim pMxDoc As IMxDocument = m_application.Document

        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pLayer As IFeatureLayer
        pLayer = pMxDoc.SelectedLayer

        If pLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a RIVER feature class in the TOC", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim pFeatClass As IFeatureClass
        pFeatClass = pLayer.FeatureClass

        If Not pFeatClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then
            MsgBox("Select a RIVER feature class in the TOC", vbCritical, "Incompatible input layer")
            Exit Sub
        End If


        For i = 0 To pFeatClass.FeatureCount(Nothing) - 1
            Dim p1RiverFeat As IFeature
            On Error Resume Next
            p1RiverFeat = pFeatClass.GetFeature(i)
            Dim p1RiverLine As IPolyline
            p1RiverLine = p1RiverFeat.Shape

            Dim pRelOp As IRelationalOperator
            pRelOp = p1RiverLine

            For j = 0 To pFeatClass.FeatureCount(Nothing) - 1
                If i <> j Then
                    Dim p2RiverFeat As IFeature
                    On Error Resume Next
                    p2RiverFeat = pFeatClass.GetFeature(j)
                    Dim pPolyline As IPolyline
                    pPolyline = p2RiverFeat.Shape
                    'Dim blnContains As Boolean = pRelOp.Contains
                    If (pRelOp.Contains(pPolyline) = True) Then
                        p2RiverFeat.Delete()
                    End If
                End If
            Next
        Next
    End Sub
End Class



