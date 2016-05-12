Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry

<ComClass(cmd_4_2_EverIncreasingHeight.ClassId, cmd_4_2_EverIncreasingHeight.InterfaceId, cmd_4_2_EverIncreasingHeight.EventsId), _
 ProgId("ArcArAz.cmd_4_2_EverIncreasingHeight")> _
Public NotInheritable Class cmd_4_2_EverIncreasingHeight
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "42abd1b2-e50d-481f-828d-6981ee901453"
    Public Const InterfaceId As String = "8992c3a9-666b-4d6b-bbe1-1bbb4c3c6cfe"
    Public Const EventsId As String = "3ceeafed-7cef-4319-bba9-b7de0ef6694d"
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
        MyBase.m_caption = "Ever Increasing Height Rivers"   'localizable text 
        MyBase.m_message = "Change negative slope to horizontal"   'localizable text 
        MyBase.m_toolTip = "Change negative slope to horizontal" 'localizable text 
        MyBase.m_name = "ArcArAz-Rivers_EverIncreasingHeight"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_4_2_EverIncreasingHeight.OnClick implementation

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pInLayer As IFeatureLayer = pMxDoc.SelectedLayer

        If pInLayer Is Nothing Then  'Check if no input layer is selected
            MsgBox("Select a RIVER feature class in the TOC", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim pInFeatClass As IFeatureClass = pInLayer.FeatureClass

        If Not pInFeatClass.ShapeType = ESRI.ArcGIS.Geometry.esriGeometryType.esriGeometryPolyline Then
            MsgBox("Select a RIVER feature class in the TOC", vbCritical, "Incompatible input layer")
            Exit Sub
        End If

        Dim pInFCursor As IFeatureCursor
        pInFCursor = pInFeatClass.Search(Nothing, True)

        Dim pFeat As IFeature
        pFeat = pInFCursor.NextFeature

        Do While Not pFeat Is Nothing
            Dim pPolyline As IPolyline = pFeat.Shape
            Dim pZAware As IZAware = pPolyline
            pZAware.ZAware = True

            If pPolyline.FromPoint.Z > pPolyline.ToPoint.Z Then
                pPolyline.ReverseOrientation()
                pFeat.Store()
            End If

            Dim slope As Double = (pPolyline.ToPoint.Z - pPolyline.FromPoint.Z) / pPolyline.Length

            Dim pPointColl As IPointCollection = pFeat.Shape
            For i = 1 To pPointColl.PointCount - 2

                Dim pPoint As IPoint = pPointColl.Point(i)
                Dim pZAwareP As IZAware = pPoint
                pZAwareP.ZAware = True
                Dim distAlong As Double = GetDistanceAt(pPoint, pPolyline)
                pPoint.Z = pPolyline.FromPoint.Z + slope * distAlong
                pPointColl.UpdatePoint(i, pPoint)

            Next
            pFeat.Shape = pPointColl
            pFeat.Store()
            pFeat = pInFCursor.NextFeature
        Loop

        '''''' Para quitar los hoyos desde arriba
        'If pPolyline.FromPoint.Z < pPolyline.ToPoint.Z Then
        '    pPolyline.ReverseOrientation()
        '    pFeat.Store()
        'End If

        'Dim pPointColl As IPointCollection = pFeat.Shape
        'For i = 0 To pPointColl.PointCount - 2
        '    If pPointColl.Point(i).Z < pPointColl.Point(i + 1).Z Then
        '        Dim pPoint As IPoint = pPointColl.Point(i + 1)
        '        Dim pZAwareP As IZAware = pPoint
        '        pZAwareP.ZAware = True
        '        pPoint.Z = pPointColl.Point(i).Z
        '        pPointColl.UpdatePoint(i + 1, pPoint)

        '    End If
        'Next


        '''''' Para quitar los hoyos desde abajo
        'If pPolyline.FromPoint.Z > pPolyline.ToPoint.Z Then
        '    pPolyline.ReverseOrientation()
        '    pFeat.Store()
        'End If

        'Dim pPointColl As IPointCollection = pFeat.Shape
        'For i = 0 To pPointColl.PointCount - 2
        '    If pPointColl.Point(i).Z > pPointColl.Point(i + 1).Z Then
        '        Dim pPoint As IPoint = pPointColl.Point(i)
        '        Dim pZAwareP As IZAware = pPoint
        '        pZAwareP.ZAware = True
        '        pPoint.Z = pPointColl.Point(i).Z
        '        pPointColl.UpdatePoint(i + 1, pPoint)

        '    End If
        'Next

    End Sub
    Public Function GetDistanceAt(ByVal point As IPoint, ByVal polyline As IPolyline)
        Dim outPnt As IPoint = New PointClass()
        Dim distAlong As Double = Nothing
        Dim distFrom As Double = Nothing
        Dim bRight As Boolean = False
        polyline.QueryPointAndDistance(esriSegmentExtension.esriNoExtension, point, False, outPnt, distAlong, distFrom, bRight)
        Return distAlong
    End Function

End Class



