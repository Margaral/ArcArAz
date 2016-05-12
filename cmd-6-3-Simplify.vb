Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.GeoprocessingUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry

<ComClass(cmd_6_3_Simplify.ClassId, cmd_6_3_Simplify.InterfaceId, cmd_6_3_Simplify.EventsId), _
 ProgId("ArcArAz.cmd_6_3_Simplify")> _
Public NotInheritable Class cmd_6_3_Simplify
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "8c7a3b77-0e21-4009-b9b6-591350f0b4ff"
    Public Const InterfaceId As String = "ffaf6cd1-c38a-4ba1-b750-4610e3385c0c"
    Public Const EventsId As String = "d7a167de-2dc3-4eb7-9771-d79ac1cd8e4d"
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
        MyBase.m_category = "ArcArAz-Vertex"  'localizable text 
        MyBase.m_caption = "Reduce vertex number"   'localizable text 
        MyBase.m_message = "Reduce the vertex number of selected shapefile"   'localizable text 
        MyBase.m_toolTip = "Reduce the vertex number of selected shapefile" 'localizable text 
        MyBase.m_name = "ArcArAz-Vertex_SimplifyCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_6_3_Simplify.OnClick implementation
        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pInLayer As IFeatureLayer = pMxDoc.SelectedLayer

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a POLYGON or POLYLINE feature class in the TOC to reduce the vertex number", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        'Set a reference to the IGPCommandHelper2 interface.
        Dim pToolHelper As IGPToolCommandHelper2 = New GPToolCommandHelper

        'Set the tool you want to invoke.
        Dim toolboxPath = "C:\Program Files (x86)\ArcGIS\Desktop10.0\ArcToolbox\Toolboxes\Cartography Tools.tbx"

        Dim pInFeatClass As IFeatureClass = pInLayer.FeatureClass

        If pInFeatClass.ShapeType <> esriGeometryType.esriGeometryPolygon And pInFeatClass.ShapeType <> esriGeometryType.esriGeometryPolyline Then
            MsgBox("Select a POLYGON or POLYLINE feature class in the TOC to reduce the vertex number", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        If pInFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
            Try
                pToolHelper.SetToolByName(toolboxPath, "SimplifyPolygon")

                'Create the messages object to pass to the InvokeModal method.
                Dim msgs As IGPMessages
                msgs = New GPMessages

                'Invoke the tool.
                pToolHelper.InvokeModal(0, Nothing, True, msgs)
                m_application.CurrentTool = Nothing
            Catch err As Exception
                MsgBox("The tool instalation folder is not the default." & vbCrLf & _
                        "Please, access this tool throught ArcToolBox: " & vbCrLf & _
                        "ArcToolBox / Cartography Tools / Generalization / Simplify Polygon.")
            End Try
            Exit Sub
        End If

        If pInFeatClass.ShapeType = esriGeometryType.esriGeometryPolyline Then
            Try
                pToolHelper.SetToolByName(toolboxPath, "SimplifyLine")

                'Create the messages object to pass to the InvokeModal method.
                Dim msgs As IGPMessages
                msgs = New GPMessages

                'Invoke the tool.
                pToolHelper.InvokeModal(0, Nothing, True, msgs)
                m_application.CurrentTool = Nothing
            Catch err As Exception
                MsgBox("The tool instalation folder is not the default." & vbCrLf & _
                        "Please, access this tool throught ArcToolBox: " & vbCrLf & _
                        "ArcToolBox / Cartography Tools / Generalization / Simplify Line.")
            End Try
            Exit Sub
        End If

    End Sub
End Class



