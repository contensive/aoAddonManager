
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.Addons.AddonManager51.Models
Imports Contensive.BaseClasses
Imports ICSharpCode.SharpZipLib

Namespace Contensive.Addons.AddonManager51
    Public Class LegacyDataRecordNodeListController
        '
        '====================================================================================================
        '
        Public Shared Function getNodeList(cp As CPBaseClass, DataRecordList As String, tempPathFileList As List(Of String), tempExportPath As String) As String
            Try
                Dim result As String = ""
                If DataRecordList <> "" Then
                    result &= vbCrLf & vbTab & "<DataRecordList>" & LegacyExportController.EncodeCData(cp, DataRecordList) & "</DataRecordList>"
                    Dim DataRecords() As String = Split(DataRecordList, vbCrLf)
                    Dim RecordNodes As String = ""
                    For Ptr = 0 To UBound(DataRecords)
                        Dim FieldNodes As String = ""
                        Dim DataRecordName As String = ""
                        Dim DataRecordGuid As String = ""
                        Dim DataRecord As String = DataRecords(Ptr)
                        If DataRecord <> "" Then
                            Dim DataSplit() As String = Split(DataRecord, ",")
                            If UBound(DataSplit) >= 0 Then
                                Dim DataContentName As String = Trim(DataSplit(0))
                                Dim DataContentId As Integer = cp.Content.GetID(DataContentName)
                                If DataContentId <= 0 Then
                                    RecordNodes = "" _
                                            & RecordNodes _
                                            & vbCrLf & vbTab & "<!-- data missing, content not found during export, content=""" & DataContentName & """ guid=""" & DataRecordGuid & """ name=""" & DataRecordName & """ -->"
                                Else
                                    Dim supportsGuid As Boolean = cp.Content.IsField(DataContentName, "ccguid")
                                    Dim Criteria As String
                                    If UBound(DataSplit) = 0 Then
                                        Criteria = ""
                                    Else
                                        Dim TestString As String = Trim(DataSplit(1))
                                        If TestString = "" Then
                                            '
                                            ' blank is a select all
                                            '
                                            Criteria = ""
                                            DataRecordName = ""
                                            DataRecordGuid = ""
                                        ElseIf Not supportsGuid Then
                                            '
                                            ' if no guid, this is name
                                            '
                                            DataRecordName = TestString
                                            DataRecordGuid = ""
                                            Criteria = "name=" & cp.Db.EncodeSQLText(DataRecordName)
                                        ElseIf (Len(TestString) = 38) And (Left(TestString, 1) = "{") And (Right(TestString, 1) = "}") Then
                                            '
                                            ' guid {726ED098-5A9E-49A9-8840-767A74F41D01} format
                                            '
                                            DataRecordGuid = TestString
                                            DataRecordName = ""
                                            Criteria = "ccguid=" & cp.Db.EncodeSQLText(DataRecordGuid)
                                        ElseIf (Len(TestString) = 36) And (Mid(TestString, 9, 1) = "-") Then
                                            '
                                            ' guid 726ED098-5A9E-49A9-8840-767A74F41D01 format
                                            '
                                            DataRecordGuid = TestString
                                            DataRecordName = ""
                                            Criteria = "ccguid=" & cp.Db.EncodeSQLText(DataRecordGuid)
                                        ElseIf (Len(TestString) = 32) And (InStr(1, TestString, " ") = 0) Then
                                            '
                                            ' guid 726ED0985A9E49A98840767A74F41D01 format
                                            '
                                            DataRecordGuid = TestString
                                            DataRecordName = ""
                                            Criteria = "ccguid=" & cp.Db.EncodeSQLText(DataRecordGuid)
                                        Else
                                            '
                                            ' use name
                                            '
                                            DataRecordName = TestString
                                            DataRecordGuid = ""
                                            Criteria = "name=" & cp.Db.EncodeSQLText(DataRecordName)
                                        End If
                                    End If
                                    Using CSData As CPCSBaseClass = cp.CSNew()
                                        If Not CSData.Open(DataContentName, Criteria, "id") Then
                                            RecordNodes = "" _
                                                    & RecordNodes _
                                                    & vbCrLf & vbTab & "<!-- data missing, record not found during export, content=""" & DataContentName & """ guid=""" & DataRecordGuid & """ name=""" & DataRecordName & """ -->"
                                        Else
                                            '
                                            ' determine all valid fields
                                            '
                                            Dim fieldCnt As Integer = 0
                                            Dim Sql As String = "select * from ccFields where contentid=" & DataContentId

                                            Dim fieldLookupListValue As String = ""
                                            Dim fieldNames() As String = {}
                                            Dim fieldTypes() As Integer = {}
                                            Dim fieldLookupContent() As String = {}
                                            Dim fieldLookupList() As String = {}
                                            Dim FieldLookupContentName As String
                                            Dim FieldTypeNumber As Integer
                                            Dim FieldName As String
                                            Using csFields As CPCSBaseClass = cp.CSNew()
                                                If csFields.Open("content fields", "contentid=" & DataContentId) Then
                                                    Do
                                                        FieldName = csFields.GetText("name")
                                                        If FieldName <> "" Then
                                                            Dim FieldLookupContentID As Integer = 0
                                                            FieldLookupContentName = ""
                                                            FieldTypeNumber = csFields.GetInteger("type")
                                                            Select Case LCase(FieldName)
                                                                Case "ccguid", "name", "id", "dateadded", "createdby", "modifiedby", "modifieddate", "createkey", "contentcontrolid", "editsourceid", "editarchive", "editblank", "contentcategoryid"
                                                                Case Else
                                                                    If FieldTypeNumber = 7 Then
                                                                        FieldLookupContentID = csFields.GetInteger("Lookupcontentid")
                                                                        fieldLookupListValue = csFields.GetText("LookupList")
                                                                        If FieldLookupContentID <> 0 Then
                                                                            FieldLookupContentName = cp.Content.GetRecordName("content", FieldLookupContentID)
                                                                        End If
                                                                    End If
                                                                    Select Case FieldTypeNumber
                                                                        Case FieldTypeFile, FieldTypeImage, FieldTypeLookup, FieldTypeBoolean, FieldTypeCSSFile, FieldTypeJavascriptFile, FieldTypeTextFile, FieldTypeXMLFile, FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger, FieldTypeDate, FieldTypeLink, FieldTypeLongText, FieldTypeResourceLink, FieldTypeText, FieldTypeHTML, FieldTypeHTMLFile
                                                                            '
                                                                            ' this is a keeper
                                                                            '
                                                                            ReDim Preserve fieldNames(fieldCnt)
                                                                            ReDim Preserve fieldTypes(fieldCnt)
                                                                            ReDim Preserve fieldLookupContent(fieldCnt)
                                                                            ReDim Preserve fieldLookupList(fieldCnt)
                                                                            'fieldLookupContent
                                                                            fieldNames(fieldCnt) = FieldName
                                                                            fieldTypes(fieldCnt) = FieldTypeNumber
                                                                            fieldLookupContent(fieldCnt) = FieldLookupContentName
                                                                            fieldLookupList(fieldCnt) = fieldLookupListValue
                                                                            fieldCnt = fieldCnt + 1
                                                                            'end case
                                                                    End Select
                                                                    'end case
                                                            End Select
                                                        End If
                                                        Call csFields.GoNext()
                                                    Loop While csFields.OK()
                                                End If
                                                Call csFields.Close()
                                            End Using
                                            '
                                            ' output records
                                            '
                                            DataRecordGuid = ""
                                            Do While CSData.OK()
                                                FieldNodes = ""
                                                DataRecordName = CSData.GetText("name")
                                                If supportsGuid Then
                                                    DataRecordGuid = CSData.GetText("ccguid")
                                                    If DataRecordGuid = "" Then
                                                        DataRecordGuid = cp.Utils.CreateGuid()
                                                        Call CSData.SetField("ccGuid", DataRecordGuid)
                                                    End If
                                                End If
                                                Dim fieldPtr As Integer
                                                For fieldPtr = 0 To fieldCnt - 1
                                                    FieldName = fieldNames(fieldPtr)
                                                    FieldTypeNumber = cp.Utils.EncodeInteger(fieldTypes(fieldPtr))
                                                    'Dim ContentID As Integer
                                                    Dim FieldValue As String
                                                    Select Case FieldTypeNumber
                                                        Case FieldTypeFile, FieldTypeImage
                                                            '
                                                            ' files -- copy pathFilename to tmp folder and save pathFilename to fieldValue
                                                            FieldValue = CSData.GetText(FieldName).ToString()
                                                            If (Not String.IsNullOrWhiteSpace(FieldValue)) Then
                                                                Dim pathFilename As String = FieldValue
                                                                cp.CdnFiles.Copy(pathFilename, tempExportPath & pathFilename, cp.TempFiles)
                                                                If Not tempPathFileList.Contains(tempExportPath & pathFilename) Then
                                                                    tempPathFileList.Add(tempExportPath & pathFilename)
                                                                    Dim path As String = genericController.getPath(pathFilename)
                                                                    Dim filename As String = genericController.getFilename(pathFilename)
                                                                    result &= vbCrLf & vbTab & "<Resource name=""" & System.Net.WebUtility.HtmlEncode(filename) & """ type=""content"" path=""" & System.Net.WebUtility.HtmlEncode(path) & """ />"
                                                                End If
                                                            End If
                                                        Case FieldTypeBoolean
                                                            '
                                                            ' true/false
                                                            '
                                                            FieldValue = CSData.GetBoolean(FieldName).ToString()
                                                        Case FieldTypeCSSFile, FieldTypeJavascriptFile, FieldTypeTextFile, FieldTypeXMLFile
                                                            '
                                                            ' text files
                                                            '
                                                            FieldValue = CSData.GetText(FieldName)
                                                            FieldValue = LegacyExportController.EncodeCData(cp, FieldValue)
                                                        Case FieldTypeInteger
                                                            '
                                                            ' integer
                                                            '
                                                            FieldValue = CSData.GetInteger(FieldName).ToString()
                                                        Case FieldTypeCurrency, FieldTypeFloat
                                                            '
                                                            ' numbers
                                                            '
                                                            FieldValue = CSData.GetNumber(FieldName).ToString()
                                                        Case FieldTypeDate
                                                            '
                                                            ' date
                                                            '
                                                            FieldValue = CSData.GetDate(FieldName).ToString()
                                                        Case FieldTypeLookup
                                                            '
                                                            ' lookup
                                                            '
                                                            FieldValue = ""
                                                            Dim FieldValueInteger As Integer = CSData.GetInteger(FieldName)
                                                            If (FieldValueInteger <> 0) Then
                                                                FieldLookupContentName = fieldLookupContent(fieldPtr)
                                                                fieldLookupListValue = fieldLookupList(fieldPtr)
                                                                If (FieldLookupContentName <> "") Then
                                                                    '
                                                                    ' content lookup
                                                                    '
                                                                    If cp.Content.IsField(FieldLookupContentName, "ccguid") Then
                                                                        Using CSlookup As CPCSBaseClass = cp.CSNew()
                                                                            Call CSlookup.OpenRecord(FieldLookupContentName, FieldValueInteger)
                                                                            If CSlookup.OK() Then
                                                                                FieldValue = CSlookup.GetText("ccguid")
                                                                                If FieldValue = "" Then
                                                                                    FieldValue = cp.Utils.CreateGuid()
                                                                                    Call CSlookup.SetField("ccGuid", FieldValue)
                                                                                End If
                                                                            End If
                                                                            Call CSlookup.Close()
                                                                        End Using
                                                                    End If
                                                                ElseIf fieldLookupListValue <> "" Then
                                                                    '
                                                                    ' list lookup, ok to save integer
                                                                    '
                                                                    FieldValue = FieldValueInteger.ToString()
                                                                End If
                                                            End If
                                                        Case Else
                                                            '
                                                            ' text types
                                                            '
                                                            FieldValue = CSData.GetText(FieldName)
                                                            FieldValue = LegacyExportController.EncodeCData(cp, FieldValue)
                                                    End Select
                                                    FieldNodes = FieldNodes & vbCrLf & vbTab & "<field name=""" & System.Net.WebUtility.HtmlEncode(FieldName) & """>" & FieldValue & "</field>"
                                                Next
                                                RecordNodes = "" _
                                                        & RecordNodes _
                                                        & vbCrLf & vbTab & "<record content=""" & System.Net.WebUtility.HtmlEncode(DataContentName) & """ guid=""" & DataRecordGuid & """ name=""" & System.Net.WebUtility.HtmlEncode(DataRecordName) & """>" _
                                                        & LegacyExportController.tabIndent(cp, FieldNodes) _
                                                        & vbCrLf & vbTab & "</record>"
                                                Call CSData.GoNext()
                                            Loop
                                        End If
                                        Call CSData.Close()
                                    End Using
                                End If
                            End If
                        End If
                    Next
                    If RecordNodes <> "" Then
                        result = "" _
                                & result _
                                & vbCrLf & vbTab & "<data>" _
                                & LegacyExportController.tabIndent(cp, RecordNodes) _
                                & vbCrLf & vbTab & "</data>"
                    End If
                End If
                Return result
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Return String.Empty
            End Try
        End Function
    End Class
End Namespace