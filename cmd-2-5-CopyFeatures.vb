Imports System.Runtime.InteropServices
Imports System.Drawing
Imports ESRI.ArcGIS.ADF.BaseClasses
Imports ESRI.ArcGIS.ADF.CATIDs
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geodatabase

<ComClass(cmd_2_5_CopyFeatures.ClassId, cmd_2_5_CopyFeatures.InterfaceId, cmd_2_5_CopyFeatures.EventsId), _
 ProgId("ArcArAz.cmd_2_5_CopyFeatures")> _
Public NotInheritable Class cmd_2_5_CopyFeatures
    Inherits BaseCommand

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "d10aaab6-dac7-4945-a74d-f339280a9a29"
    Public Const InterfaceId As String = "1a26e71b-6d4a-4695-84ed-0ff01966feed"
    Public Const EventsId As String = "8a0d21d4-4300-4979-9de2-52555735e995"
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
        MyBase.m_caption = "Copy/Paste features"   'localizable text 
        MyBase.m_message = "Copy features in the same FC"   'localizable text 
        MyBase.m_toolTip = "Copy features in the same FC" 'localizable text 
        MyBase.m_name = "ArcArAz-EntityProp_CopyFeature"  'unique id, non-localizable (e.g. "MyCategory_ArcMapCommand")

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
        'TODO: Add cmd_2_5_CopyFeatures.OnClick implementation

        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap

        Dim i As Integer

        For i = 0 To pMap.LayerCount - 1
            Dim j As Integer = 0
            If TypeOf pMap.Layer(i) Is IFeatureLayer Then

                Dim pFeatSel As IFeatureSelection = pMap.Layer(i)
                Dim pSelSet As ISelectionSet = pFeatSel.SelectionSet
                Dim pInFCursor As IFeatureCursor = Nothing

                If pSelSet.Count <> 0 Then

                    pFeatSel.SelectionSet.Search(Nothing, True, pInFCursor)
                    Dim pFeatLayer As IFeatureLayer = pMap.Layer(i)
                    Dim pFeatClass As IFeatureClass = pFeatLayer.FeatureClass
                    Dim pDataSet As IDataset = pFeatClass
                    Dim pWSE As IWorkspaceEdit = pDataSet.Workspace
                    pWSE.StartEditing(False)

                    Dim pFeat As IFeature = pInFCursor.NextFeature

                    For j = 0 To pSelSet.Count - 1

                        Dim pFields As IFields = pFeatClass.Fields
                        Dim pFeat2 As IFeature = pFeatClass.CreateFeature
                        pFeat2.Shape = pFeat.Shape

                        For k = 0 To pFields.FieldCount - 1
                            Dim pField As IField = pFields.Field(k)
                            If pField.Editable = True Then
                                pFeat2.Value(k) = pFeat.Value(k)
                            End If
                        Next
                        pFeat2.Store()
                        pFeat = pInFCursor.NextFeature
                    Next
                    MsgBox(j & " features have been copied into " & pMap.Layer(i).Name)
                    pWSE.StopEditing(True)
                End If
            End If

        Next


    End Sub

    '    Public Overrides Sub OnClick()
    '        'TODO: Add cmd_2_5_CopyFeatures.OnClick implementation

    '        Dim pMxDoc As IMxDocument = m_application.Document
    '        Dim pMap As IMap = pMxDoc.FocusMap

    '        Dim i As Integer

    '        For i = 0 To pMap.LayerCount - 1
    '            Dim j As Integer = 0
    '            If TypeOf pMap.Layer(i) Is IFeatureLayer Then
    '                Dim pFeatSel As IFeatureSelection = pMap.Layer(i)
    '                Dim pSelSet As ISelectionSet = pFeatSel.SelectionSet
    '                Dim pInFCursor As IFeatureCursor = Nothing

    '                If pSelSet.Count <> 0 Then
    '                    pFeatSel.SelectionSet.Search(Nothing, True, pInFCursor)
    '                    Dim pFeatLayer As IFeatureLayer = pMap.Layer(i)
    '                    Dim pFeatClass As IFeatureClass = pFeatLayer.FeatureClass

    '                    Dim featureBuffer As IFeatureBuffer = pFeatClass.CreateFeatureBuffer()
    '                    Dim featureCursor As IFeatureCursor = pFeatClass.Insert(True)


    '                    Dim pFeat As IFeature = pInFCursor.NextFeature


    '                    For j = 0 To pSelSet.Count - 1

    '                        Dim pFields As IFields = pFeatClass.Fields
    '                        'Dim pFeat2 As IFeature = pFeatClass.CreateFeature
    '                        'pFeat2.Shape = pFeat.Shape

    '                        For k = 0 To pFields.FieldCount - 1
    '                            Dim pField As IField = pFields.Field(k)
    '                            If pField.Editable = True Then
    '                                featureBuffer.Value(k) = pFeat.Value(k)
    '                                'pFeat2.Value(k) = pFeat.Value(k)
    '                            End If
    '                        Next
    '                        featureBuffer.Shape = pFeat.Shape
    '                        featureCursor.InsertFeature(featureBuffer)
    '                        'pFeat2.Store()
    '                        pFeat = pInFCursor.NextFeature

    '                    Next
    '                    MsgBox(j & " features have been copied into " & pMap.Layer(i).Name)
    '                    Try
    '                        featureCursor.Flush()
    '                    Catch comExc As COMException
    '                        ' Handle the error in a way appropriate to your application.
    '                    Finally
    '                        ' Release the cursor as it's no longer needed.
    '                        Marshal.ReleaseComObject(featureCursor)
    '                    End Try
    '                End If
    '            End If

    '        Next

    '    End Sub
End Class



