Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.DataSourcesFile
Imports System.IO
Imports ESRI.ArcGIS.ArcMapUI

Public Class cmd_2_4_Frm_1_SelectFieldName

    Private m_application As IApplication

    Public Property Application() As IApplication
        Get
            Return m_application
        End Get
        Set(ByVal value As IApplication)
            m_application = value
        End Set
    End Property

    Private Sub cmd_2_4_Frm_1_SelectFieldName_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim pMxDoc As IMxDocument = m_application.Document

        Dim pMap As IMap = pMxDoc.FocusMap

        Dim pInLayer As IFeatureLayer = pMxDoc.SelectedLayer

        Dim pInFeatClass As IFeatureClass = pInLayer.FeatureClass

        Dim pFields As IFields = pInFeatClass.Fields

        Dim i As Integer
        For i = 0 To pFields.FieldCount - 1
            Me.ComboBox1.Items.Add(pFields.Field(i).Name)
        Next

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim strNewNameField As String = InputBox("Set new field's name")
        Dim pMxDoc As IMxDocument = m_application.Document
        Dim pMap As IMap = pMxDoc.FocusMap
        Dim pInLayer As IFeatureLayer = pMxDoc.SelectedLayer
        Dim pInFeatClass As IFeatureClass = pInLayer.FeatureClass

        Dim pFieldEdit As IFieldEdit = New Field

        pFieldEdit.Name_2 = strNewNameField
        pFieldEdit.Type_2 = esriFieldType.esriFieldTypeString

        pFieldEdit.Length_2 = 10
        pInFeatClass.AddField(pFieldEdit)

        Dim pFields As IFields
        pFields = pInFeatClass.Fields

        Dim lngOldF As Long
        lngOldF = pFields.FindField(ComboBox1.Text)

        Dim lngNewF As Long
        lngNewF = pFields.FindField(strNewNameField)

        Dim pTable As ITable
        pTable = pInFeatClass

        Dim pCur As ICursor
        pCur = pTable.Search(Nothing, False)

        Dim pRow As IRow
        pRow = pCur.NextRow

        Do While Not pRow Is Nothing
            Dim strName As String
            strName = UCase(pRow.Value(lngOldF))
            strName = Replace(strName, "Ñ", "N")
            strName = Replace(strName, "DE ", "")
            strName = Replace(strName, "DEL ", "")
            strName = Replace(strName, "LA ", "")
            strName = Replace(strName, "LAS ", "")
            strName = Replace(strName, "EL ", "")
            strName = Replace(strName, "LOS ", "")
            strName = Replace(strName, " ", "")
            strName = Replace(strName, "Á", "A")
            strName = Replace(strName, "É", "E")
            strName = Replace(strName, "Í", "I")
            strName = Replace(strName, "Ó", "O")
            strName = Replace(strName, "Ú", "U")
            strName = Replace(strName, "Ü", "U")
            '), " ", "_"), "Ñ", "N"), "DE", ""),"DEL",""),"LA","")

            pRow.Value(lngNewF) = Microsoft.VisualBasic.Left(strName, 7) & pRow.Value(0)
            pRow.Store()
            pRow = pCur.NextRow

        Loop
        Me.Hide()
    End Sub
End Class