Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

<ComClass(cmd2_1RemoveInteriorHoles.ClassId, cmd2_1RemoveInteriorHoles.InterfaceId, cmd2_1RemoveInteriorHoles.EventsId), _
 ProgId("ArcArAz.cmd2_1RemoveInteriorHoles")> _
Public NotInheritable Class cmd2_1RemoveInteriorHoles
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "8bac7d2f-79bd-4453-be92-1a1af544e9b3"
    Public Const InterfaceId As String = "0f9680cb-13d7-40fc-8ce4-e0327207f820"
    Public Const EventsId As String = "afc38d1a-260e-4e7c-a0ea-11f57210a8e9"
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
        MyBase.m_category = "ArcArAz-EntityProp"  'localizable text 
        MyBase.m_caption = "Remove Interior Holes"   'localizable text 
        MyBase.m_message = "Remove holes from polygons"   'localizable text 
        MyBase.m_toolTip = "Remove holes from polygons" 'localizable text 
        MyBase.m_name = "ArcArAz-Input_RemoveInteriorHolesCommand"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd2_1RemoveInteriorHoles.OnClick implementation
        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pInLayer As IFeatureLayer = pMxDoc.SelectedLayer

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a POLYGON feature class in the TOC to check existing holes and remove them", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim pInFeatClass As IFeatureClass = pInLayer.FeatureClass

        If Not pInFeatClass.ShapeType = esriGeometryType.esriGeometryPolygon Then
            MsgBox("Select a POLYGON feature class in the TOC to check existing holes and remove them", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim i As Integer
        Dim pFeat As IFeature
        Dim pPolygon As IPolygon
        Dim pRings As IGeometryCollection
        Dim pTopo As ITopologicalOperator2
        Dim pPart As IArea
        Dim bHoles As Boolean
        Dim j As Integer
        Dim k As Integer = 0

        For i = 0 To pInFeatClass.FeatureCount(Nothing) - 1
            pFeat = pInFeatClass.GetFeature(i)
            pPolygon = pFeat.Shape
            pRings = pPolygon
            pTopo = pPolygon
            'pTopo.IsKnownSimple_1 = False

            bHoles = True
            Do While bHoles
                bHoles = False
                pPolygon.SimplifyPreserveFromTo()
                For j = pRings.GeometryCount - 1 To 0 Step -1
                    pPart = pRings.Geometry(j)
                    If pPart.Area <= 0 Then
                        pRings.RemoveGeometries(j, 1)
                        bHoles = True
                        k = k + 1
                    End If
                Next
            Loop
            pFeat.Shape = pPolygon
            pFeat.Store()
        Next

        MsgBox("Total removed holes: " & k)

        pMxDoc.ActiveView.ContentsChanged()
        pMxDoc.UpdateContents()

    End Sub
End Class



