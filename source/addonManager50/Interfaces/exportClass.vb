

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports ICSharpCode.SharpZipLib

Namespace Contensive.Addons.AddonManager
    Public Class exportClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        ''' <summary>
        ''' export collection
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Dim returnResult As String = ""
                Try
                    Dim Button As String = CP.Doc.GetText(RequestNameButton)
                    If Button = ButtonCancel Then
                        '
                        ' ----- redirect back to the root
                        Call CP.Response.Redirect(CP.Site.GetText("adminUrl"))
                    Else
                        '
                        ' -- create form
                        Dim form As New adminFramework.formNameValueRowsClass
                        form.title = "Export Collection"
                        form.body = CP.Html.p("Use this tool to create an Add-on Collection zip file that can be used to install a collection on another site.")
                        If Not CP.User.IsAdmin() Then
                            '
                            ' -- Put up error message
                            form.body &= CP.Html.p("You must be an administrator to use this tool.")
                        Else
                            If (CP.Site.GetText("buildVersion") < CP.Version) Then
                                '
                                ' -- database needs to be upgraded
                                form.description &= CP.Html.p("Warning: The site's Database needs to be upgraded. You should do this before installing addons.")
                            End If
                            '
                            ' -- Upload tool
                            If (Button = ButtonOK) Then
                                    '
                                    ' -- export
                                    Dim CollectionID As Integer = CP.Doc.GetInteger(RequestNameCollectionID)
                                    Dim addonCollection As Models.addonCollectionModel = Models.addonCollectionModel.create(CP, CollectionID)
                                    If (addonCollection Is Nothing) Then
                                        '
                                        ' -- collection not found
                                        Call CP.UserError.Add("The collection file you selected could not be found. Please select another.")
                                    Else
                                        '
                                        ' -- build collection zip file and return file
                                        Dim CollectionFilename As String = GetCollectionZipPathFilename(CP, CollectionID)
                                        If Not CP.UserError.OK Then
                                            '
                                            ' -- errors during export
                                            form.body = CP.Html.div(CP.Html.p("ERRORS during export: ") & CP.Html.ul(CP.UserError.GetList()))
                                        Else
                                            '
                                            ' -- success
                                            form.body &= CP.Html.p("Export Successful")
                                            form.body &= CP.Html.p("Click <a href=""" & CP.Site.FilePath & Replace(CollectionFilename, "\", "/") & """>here</a> to download the collection file</p>")
                                        End If
                                    End If
                                End If
                                '
                                ' Get Form
                                form.addRow()
                                form.rowName = "Add-on Collection"
                                form.rowValue = CP.Html.SelectContent(RequestNameCollectionID, "0", "Add-on Collections")
                                form.addFormHidden("UploadCount", "1")
                                form.addFormButton(ButtonOK)
                                form.addFormButton(ButtonCancel)
                            End If
                            returnResult = form.getHtml(CP)
                    End If
                Catch ex As Exception
                    CP.Site.ErrorReport(ex)
                End Try
                Return returnResult
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
            End Try
            Return returnHtml
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' error handler
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="ex"></param>
        ''' <param name="method"></param>
        ''' <remarks></remarks>
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in sampleClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
        '
        '====================================================================================================
        Private Function GetCollectionZipPathFilename(cp As CPBaseClass, CollectionID As Integer) As String
            Dim collectionZipPathFilename As String = ""
            Try
                Dim IncludeSharedStyleGuidList As String
                Dim isUpdatable As Boolean
                Dim fieldLookupListValue As String = ""
                Dim FieldValueInteger As Integer
                Dim FieldLookupContentName As String
                Dim fieldPtr As Integer
                Dim fieldCnt As Integer
                Dim fieldNames() As String = {}
                Dim fieldTypes() As Integer = {}
                Dim fieldLookupContent() As String = {}
                Dim fieldLookupList() As String = {}
                Dim FieldLookupContentID As Integer
                Dim Criteria As String
                Dim supportsGuid As Boolean
                Dim reload As Boolean
                Dim ContentID As Integer
                Dim FieldValue As String
                Dim FieldTypeNumber As Integer
                Dim DataRecordList As String
                Dim DataRecords() As String
                Dim DataRecord As String
                Dim DataSplit() As String
                Dim DataContentName As String
                Dim DataContentId As Integer
                Dim DataRecordGuid As String
                Dim DataRecordName As String
                Dim TestString As String
                Dim FieldName As String
                Dim FieldNodes As String
                Dim RecordNodes As String
                Dim Modules() As String
                Dim ModuleGuid As String
                Dim Code As String
                Dim ManualFilename As String = ""
                Dim ResourceCnt As Integer
                Dim ContentName As String
                Dim FileList As String
                Dim Files() As String
                Dim Ptr As Integer
                Dim PathFilename As String
                Dim Filename As String
                Dim Path As String
                Dim Pos As Integer
                Dim collectionXml As String
                Dim Node As String
                Dim CollectionGuid As String
                Dim Guid As String
                Dim ArchiveFilename As String
                Dim ArchivePath As String
                Dim InstallFilename As String
                Dim CollectionName As String
                Dim AddFilename As String
                Dim PhysicalWWWPath As String
                Dim CollectionPath As String = ""
                Dim LastChangeDate As Date
                Dim AddonPath As String
                Dim AddFileList As New List(Of String)
                Dim IncludeModuleGuidList As String = ""
                Dim ExecFileListNode As String = ""
                Dim blockNavigatorNode As Boolean
                Dim CSData As CPCSBaseClass = cp.CSNew()
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim CS2 As CPCSBaseClass = cp.CSNew()
                Dim CS3 As CPCSBaseClass = cp.CSNew()
                Dim CSlookup As CPCSBaseClass = cp.CSNew()
                '
                IncludeSharedStyleGuidList = ""
                '
                CS.OpenRecord("Add-on Collections", CollectionID)
                If Not CS.OK() Then
                    Call cp.UserError.Add("The collection you selected could not be found")
                Else
                    CollectionGuid = CS.GetText("ccGuid")
                    If CollectionGuid = "" Then
                        CollectionGuid = cp.Utils.CreateGuid()
                        Call CS.SetField("ccGuid", CollectionGuid)
                    End If
                    CollectionName = CS.GetText("name")
                    If Not CS.FieldOK("updatable") Then
                        isUpdatable = True
                    Else
                        isUpdatable = CS.GetBoolean("updatable")
                    End If
                    If Not CS.FieldOK("blockNavigatorNode") Then
                        blockNavigatorNode = False
                    Else
                        blockNavigatorNode = CS.GetBoolean("blockNavigatorNode")
                    End If
                    collectionXml = "" _
                        & "<?xml version=""1.0"" encoding=""windows-1252""?>" _
                        & vbCrLf & "<Collection name=""" & cp.Utils.EncodeHTML(CollectionName) & """ guid=""" & CollectionGuid & """ system=""" & kmaGetYesNo(cp, CS.GetBoolean("system")) & """ updatable=""" & kmaGetYesNo(cp, isUpdatable) & """ blockNavigatorNode=""" & kmaGetYesNo(cp, blockNavigatorNode) & """>"
                    '
                    ' Archive Filenames
                    '
                    'Call Main.testpoint("getCollection, 200")
                    ArchivePath = cp.Site.PhysicalFilePath & "CollectionExport\"
                    'Call Main.testpoint("getCollection, 201")
                    InstallFilename = encodeFilename(cp, CollectionName & ".xml")
                    'Call Main.testpoint("getCollection, 202")
                    InstallFilename = ArchivePath & InstallFilename
                    'Call Main.testpoint("getCollection, 203")
                    ArchiveFilename = encodeFilename(cp, CollectionName & ".zip")
                    'Call Main.testpoint("getCollection, 204")
                    ArchiveFilename = ArchivePath & ArchiveFilename
                    'Call Main.testpoint("getCollection, 205")
                    collectionZipPathFilename = "CollectionExport\" & encodeFilename(cp, CollectionName & ".zip")
                    'Call Main.testpoint("getCollection, 207")
                    '
                    ' Delete old archive file
                    '
                    Call cp.File.Delete(ArchiveFilename)
                    '
                    '
                    ' Build executable file list Resource Node so executables can be added to addons for Version40compatibility
                    '   but save it for the end, executableFileList
                    '
                    'Call Main.testpoint("getCollection, 400")
                    AddonPath = cp.Site.PhysicalInstallPath & "\addons\"
                    FileList = CS.GetText("execFileList")
                    If FileList <> "" Then
                        '
                        ' There are executable files to include in the collection
                        '   If installed, source path is collectionpath, if not installed, collectionpath will be empty
                        '   and file will be sourced right from addon path
                        '
                        Call GetLocalCollectionArgs(cp, CollectionGuid, CollectionPath, LastChangeDate)
                        If CollectionPath <> "" Then
                            CollectionPath = CollectionPath & "\"
                        End If
                        Files = Split(FileList, vbCrLf)
                        For Ptr = 0 To UBound(Files)
                            PathFilename = Files(Ptr)
                            If PathFilename <> "" Then
                                PathFilename = Replace(PathFilename, "\", "/")
                                Path = ""
                                Filename = PathFilename
                                Pos = InStrRev(PathFilename, "/")
                                If Pos > 0 Then
                                    Filename = Mid(PathFilename, Pos + 1)
                                    Path = Mid(PathFilename, 1, Pos - 1)
                                End If
                                If LCase(Filename) <> LCase(ManualFilename) Then
                                    AddFilename = AddonPath & CollectionPath & Filename
                                    If Not AddFileList.Contains(AddFilename) Then
                                        AddFileList.Add(AddFilename)
                                        ExecFileListNode = ExecFileListNode & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""executable"" path=""" & cp.Utils.EncodeHTML(Path) & """ />"
                                    End If
                                End If
                                ResourceCnt = ResourceCnt + 1
                            End If
                        Next
                    End If
                    'Call Main.testpoint("getCollection, 500")
                    If (ResourceCnt = 0) And (CollectionPath <> "") Then
                        '
                        ' If no resources were in the collection record, this might be an old installation
                        ' Add all .dll files in the CollectionPath
                        '
                        ExecFileListNode = ExecFileListNode & AddCompatibilityResources(cp, AddonPath & CollectionPath, ArchiveFilename, "")
                    End If
                    '
                    ' helpLink
                    '
                    If CS.FieldOK("HelpLink") Then
                        collectionXml = collectionXml & vbCrLf & vbTab & "<HelpLink>" & cp.Utils.EncodeHTML(CS.GetText("HelpLink")) & "</HelpLink>"
                    End If
                    '
                    ' Help
                    '
                    collectionXml = collectionXml & vbCrLf & vbTab & "<Help>" & cp.Utils.EncodeHTML(CS.GetText("Help")) & "</Help>"
                    '
                    ' Addons
                    '
                    CS2.Open("Add-ons", "collectionid=" & CollectionID, "name", True, "id")
                    Do While CS2.OK()
                        collectionXml = collectionXml & GetAddonNode(cp, CS2.GetInteger("id"), IncludeModuleGuidList, IncludeSharedStyleGuidList)
                        Call CS2.GoNext()
                    Loop
                    '
                    ' Data Records
                    '
                    'Call Main.testpoint("getCollection, 600")
                    DataRecordList = CS.GetText("DataRecordList")
                    If DataRecordList <> "" Then
                        DataRecords = Split(DataRecordList, vbCrLf)
                        RecordNodes = ""
                        For Ptr = 0 To UBound(DataRecords)
                            FieldNodes = ""
                            DataRecordName = ""
                            DataRecordGuid = ""
                            DataRecord = DataRecords(Ptr)
                            If DataRecord <> "" Then
                                DataSplit = Split(DataRecord, ",")
                                If UBound(DataSplit) >= 0 Then
                                    DataContentName = Trim(DataSplit(0))
                                    DataContentId = cp.Content.GetID(DataContentName)
                                    If DataContentId <= 0 Then
                                        RecordNodes = "" _
                                            & RecordNodes _
                                            & vbCrLf & vbTab & "<!-- data missing, content not found during export, content=""" & DataContentName & """ guid=""" & DataRecordGuid & """ name=""" & DataRecordName & """ -->"
                                    Else
                                        supportsGuid = cp.Content.IsField(DataContentName, "ccguid")
                                        If UBound(DataSplit) = 0 Then
                                            Criteria = ""
                                        Else
                                            TestString = Trim(DataSplit(1))
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
                                        If Not CSData.Open(DataContentName, Criteria, "id") Then
                                            RecordNodes = "" _
                                                & RecordNodes _
                                                & vbCrLf & vbTab & "<!-- data missing, record not found during export, content=""" & DataContentName & """ guid=""" & DataRecordGuid & """ name=""" & DataRecordName & """ -->"
                                        Else
                                            '
                                            ' determine all valid fields
                                            '
                                            fieldCnt = 0
                                            Dim Sql As String = "select * from ccFields where contentid=" & DataContentId
                                            Dim csFields As CPCSBaseClass = cp.CSNew()
                                            If csFields.Open("content fields", "contentid=" & DataContentId) Then
                                                Do
                                                    FieldName = csFields.GetText("name")
                                                    If FieldName <> "" Then
                                                        FieldLookupContentID = 0
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
                                                                    Case FieldTypeLookup, FieldTypeBoolean, FieldTypeCSSFile, FieldTypeJavascriptFile, FieldTypeTextFile, FieldTypeXMLFile, FieldTypeCurrency, FieldTypeFloat, FieldTypeInteger, FieldTypeDate, FieldTypeLink, FieldTypeLongText, FieldTypeResourceLink, FieldTypeText, FieldTypeHTML, FieldTypeHTMLFile
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
                                                For fieldPtr = 0 To fieldCnt - 1
                                                    FieldName = fieldNames(fieldPtr)
                                                    FieldTypeNumber = cp.Utils.EncodeInteger(fieldTypes(fieldPtr))
                                                    Select Case FieldTypeNumber
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
                                                            FieldValue = EncodeCData(cp, FieldValue)
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
                                                            FieldValueInteger = CSData.GetInteger(FieldName)
                                                            If (FieldValueInteger <> 0) Then
                                                                FieldLookupContentName = fieldLookupContent(fieldPtr)
                                                                fieldLookupListValue = fieldLookupList(fieldPtr)
                                                                If (FieldLookupContentName <> "") Then
                                                                    '
                                                                    ' content lookup
                                                                    '
                                                                    If cp.Content.IsField(FieldLookupContentName, "ccguid") Then
                                                                        Call CSlookup.OpenRecord(FieldLookupContentName, FieldValueInteger)
                                                                        If CSlookup.OK() Then
                                                                            FieldValue = CSlookup.GetText("ccguid")
                                                                            If FieldValue = "" Then
                                                                                FieldValue = cp.Utils.CreateGuid()
                                                                                Call CSlookup.SetField("ccGuid", FieldValue)
                                                                            End If
                                                                        End If
                                                                        Call CSlookup.Close()
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
                                                            FieldValue = EncodeCData(cp, FieldValue)
                                                    End Select
                                                    FieldNodes = FieldNodes & vbCrLf & vbTab & "<field name=""" & cp.Utils.EncodeHTML(FieldName) & """>" & FieldValue & "</field>"
                                                Next
                                                RecordNodes = "" _
                                                    & RecordNodes _
                                                    & vbCrLf & vbTab & "<record content=""" & cp.Utils.EncodeHTML(DataContentName) & """ guid=""" & DataRecordGuid & """ name=""" & cp.Utils.EncodeHTML(DataRecordName) & """>" _
                                                    & tabIndent(cp, FieldNodes) _
                                                    & vbCrLf & vbTab & "</record>"
                                                Call CSData.GoNext()
                                            Loop
                                        End If
                                        Call CSData.Close()
                                    End If
                                End If
                            End If
                        Next
                        If RecordNodes <> "" Then
                            collectionXml = "" _
                                & collectionXml _
                                & vbCrLf & vbTab & "<data>" _
                                & tabIndent(cp, RecordNodes) _
                                & vbCrLf & vbTab & "</data>"
                        End If
                    End If
                    '
                    ' CDef
                    '
                    'Call Main.testpoint("getCollection, 700")
                    For Each content As Models.ContentModel In Models.ContentModel.createListFromCollection(cp, CollectionID)
                        reload = False
                        If (String.IsNullOrEmpty(content.ccguid)) Then
                            content.ccguid = cp.Utils.CreateGuid()
                            content.save(cp)
                            reload = True
                        End If
                        Dim xmlTool As New xmlController(cp)
                        Node = xmlTool.GetXMLContentDefinition3(content.name)
                        '
                        ' remove the <collection> top node
                        '
                        Pos = InStr(1, Node, "<cdef", vbTextCompare)
                        If Pos > 0 Then
                            Node = Mid(Node, Pos)
                            Pos = InStr(1, Node, "</cdef>", vbTextCompare)
                            If Pos > 0 Then
                                Node = Mid(Node, 1, Pos + 6)
                                collectionXml = collectionXml & vbCrLf & vbTab & Node
                            End If
                        End If
                    Next
                    '
                    ' Scripting Modules
                    '
                    'Call Main.testpoint("getCollection, 800")

                    If IncludeModuleGuidList <> "" Then
                        Modules = Split(IncludeModuleGuidList, vbCrLf)
                        For Ptr = 0 To UBound(Modules)
                            ModuleGuid = Modules(Ptr)
                            If ModuleGuid <> "" Then
                                CS2.Open("Scripting Modules", "ccguid=" & cp.Db.EncodeSQLText(ModuleGuid))
                                If CS2.OK() Then
                                    Code = Trim(CS2.GetText("code"))
                                    Code = EncodeCData(cp, Code)
                                    collectionXml = collectionXml & vbCrLf & vbTab & "<ScriptingModule Name=""" & cp.Utils.EncodeHTML(CS2.GetText("name")) & """ guid=""" & ModuleGuid & """>" & Code & "</ScriptingModule>"
                                End If
                                Call CS2.Close()
                            End If
                        Next
                    End If
                    '
                    ' shared styles
                    '
                    Dim recordGuids() As String
                    Dim recordGuid As String
                    If (IncludeSharedStyleGuidList <> "") Then
                        recordGuids = Split(IncludeSharedStyleGuidList, vbCrLf)
                        For Ptr = 0 To UBound(recordGuids)
                            recordGuid = recordGuids(Ptr)
                            If recordGuid <> "" Then
                                CS2.Open("Shared Styles", "ccguid=" & cp.Db.EncodeSQLText(recordGuid))
                                If CS2.OK() Then
                                    collectionXml = collectionXml & vbCrLf & vbTab & "<SharedStyle" _
                                        & " Name=""" & cp.Utils.EncodeHTML(CS2.GetText("name")) & """" _
                                        & " guid=""" & recordGuid & """" _
                                        & " alwaysInclude=""" & CS2.GetBoolean("alwaysInclude") & """" _
                                        & " prefix=""" & cp.Utils.EncodeHTML(CS2.GetText("prefix")) & """" _
                                        & " suffix=""" & cp.Utils.EncodeHTML(CS2.GetText("suffix")) & """" _
                                        & " sortOrder=""" & cp.Utils.EncodeHTML(CS2.GetText("sortOrder")) & """" _
                                        & ">" _
                                        & EncodeCData(cp, Trim(CS2.GetText("styleFilename"))) _
                                        & "</SharedStyle>"
                                End If
                                Call CS2.Close()
                            End If
                        Next
                    End If
                    '
                    ' Import Collections
                    '
                    Node = ""
                    If CS3.Open("Add-on Collection Parent Rules", "parentid=" & CollectionID) Then
                        Do
                            CS2.OpenRecord("Add-on Collections", CS3.GetInteger("childid"))
                            If CS2.OK() Then
                                Guid = CS2.GetText("ccGuid")
                                If Guid = "" Then
                                    Guid = cp.Utils.CreateGuid()
                                    Call CS2.SetField("ccGuid", Guid)
                                End If
                                Node = Node & vbCrLf & vbTab & "<ImportCollection name=""" & cp.Utils.EncodeHTML(CS2.GetText("name")) & """>" & Guid & "</ImportCollection>"
                            End If
                            Call CS2.Close()
                            Call CS3.GoNext()
                        Loop While CS3.OK()
                    End If
                    Call CS3.Close()
                    collectionXml = collectionXml & Node
                    '
                    ' wwwFileList
                    '
                    ResourceCnt = 0
                    FileList = CS.GetText("wwwFileList")
                    If FileList <> "" Then
                        PhysicalWWWPath = cp.Site.PhysicalWWWPath
                        If Right(PhysicalWWWPath, 1) <> "\" Then
                            PhysicalWWWPath = PhysicalWWWPath & "\"
                        End If
                        Files = Split(FileList, vbCrLf)
                        For Ptr = 0 To UBound(Files)
                            PathFilename = Files(Ptr)
                            If PathFilename <> "" Then
                                PathFilename = Replace(PathFilename, "\", "/")
                                Path = ""
                                Filename = PathFilename
                                Pos = InStrRev(PathFilename, "/")
                                If Pos > 0 Then
                                    Filename = Mid(PathFilename, Pos + 1)
                                    Path = Mid(PathFilename, 1, Pos - 1)
                                End If
                                If LCase(Filename) = "collection.hlp" Then
                                    '
                                    ' legacy file, remove it
                                    '
                                Else
                                    PathFilename = Replace(PathFilename, "/", "\")
                                    AddFilename = PhysicalWWWPath & PathFilename
                                    If AddFileList.Contains(AddFilename) Then
                                        Call cp.UserError.Add("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                                    Else
                                        AddFileList.Add(AddFilename)
                                        collectionXml = collectionXml & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""www"" path=""" & cp.Utils.EncodeHTML(Path) & """ />"
                                    End If
                                    ResourceCnt = ResourceCnt + 1
                                End If
                            End If
                        Next
                    End If
                    '
                    ' ContentFileList
                    '
                    FileList = CS.GetText("ContentFileList")
                    If FileList <> "" Then
                        Files = Split(FileList, vbCrLf)
                        For Ptr = 0 To UBound(Files)
                            PathFilename = Files(Ptr)
                            If PathFilename <> "" Then
                                PathFilename = Replace(PathFilename, "\", "/")
                                Path = ""
                                Filename = PathFilename
                                Pos = InStrRev(PathFilename, "/")
                                If Pos > 0 Then
                                    Filename = Mid(PathFilename, Pos + 1)
                                    Path = Mid(PathFilename, 1, Pos - 1)
                                End If
                                PathFilename = Replace(PathFilename, "/", "\")
                                If Left(PathFilename, 1) = "\" Then
                                    PathFilename = Mid(PathFilename, 2)
                                End If
                                AddFilename = cp.Site.PhysicalFilePath & PathFilename
                                If AddFileList.Contains(AddFilename) Then
                                    Call cp.UserError.Add("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                                Else
                                    AddFileList.Add(AddFilename)
                                    collectionXml = collectionXml & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""content"" path=""" & cp.Utils.EncodeHTML(Path) & """ />"
                                End If
                                ResourceCnt = ResourceCnt + 1
                            End If
                        Next
                    End If
                    '
                    ' ExecFileListNode
                    '
                    collectionXml = collectionXml & ExecFileListNode
                    '
                    ' Other XML
                    '
                    Dim OtherXML As String
                    OtherXML = CS.GetText("otherxml")
                    If Trim(OtherXML) <> "" Then
                        collectionXml = collectionXml & vbCrLf & OtherXML
                    End If
                    collectionXml = collectionXml & vbCrLf & "</Collection>"
                    Call CS.Close()
                    '
                    ' Save the installation file and add it to the archive
                    '
                    Call cp.File.Save(InstallFilename, collectionXml)
                    If Not AddFileList.Contains(InstallFilename) Then
                        AddFileList.Add(InstallFilename)
                    End If
                    Call zipFile(cp, ArchiveFilename, AddFileList)
                    'Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable("@" & AddFileListFilename))
                End If
            Catch ex As Exception
                errorReport(cp, ex, "GetCollection")
            End Try
            Return collectionZipPathFilename
        End Function
        '
        '====================================================================================================

        Private Function GetAddonNode(cp As CPBaseClass, addonid As Integer, ByRef Return_IncludeModuleGuidList As String, ByRef Return_IncludeSharedStyleGuidList As String) As String
            Dim result As String = ""
            Try
                '
                Dim styleId As Integer
                Dim fieldType As String
                Dim fieldTypeID As Integer
                Dim TriggerContentID As Integer
                Dim StylesTest As String
                Dim BlockEditTools As Boolean
                Dim NavType As String
                Dim Styles As String
                Dim NodeInnerText As String
                Dim IncludedAddonID As Integer
                Dim ScriptingModuleID As Integer
                Dim Guid As String
                Dim addonName As String
                Dim processRunOnce As Boolean
                Dim CS As CPCSBaseClass = cp.CSNew()
                Dim CS2 As CPCSBaseClass = cp.CSNew()
                Dim CS3 As CPCSBaseClass = cp.CSNew()
                '
                If CS.OpenRecord("Add-ons", addonid) Then
                    addonName = CS.GetText("name")
                    processRunOnce = CS.GetBoolean("ProcessRunOnce")
                    If ((LCase(addonName) = "oninstall") Or (LCase(addonName) = "_oninstall")) Then
                        processRunOnce = True
                    End If
                    '
                    ' ActiveX DLL node is being deprecated. This should be in the collection resource section
                    '
                    result &= GetNodeText(cp, "Copy", CS.GetText("Copy"))
                    result &= GetNodeText(cp, "CopyText", CS.GetText("CopyText"))
                    '
                    ' DLL
                    '

                    result &= GetNodeText(cp, "ActiveXProgramID", CS.GetText("objectprogramid"))
                    result &= GetNodeText(cp, "DotNetClass", CS.GetText("DotNetClass"))
                    '
                    ' Features
                    '
                    result &= GetNodeText(cp, "ArgumentList", CS.GetText("ArgumentList"))
                    result &= GetNodeBoolean(cp, "AsAjax", CS.GetBoolean("AsAjax"))
                    result &= GetNodeBoolean(cp, "Filter", CS.GetBoolean("Filter"))
                    result &= GetNodeText(cp, "Help", CS.GetText("Help"))
                    result &= GetNodeText(cp, "HelpLink", CS.GetText("HelpLink"))
                    result &= vbCrLf & vbTab & "<Icon Link=""" & CS.GetText("iconfilename") & """ width=""" & CS.GetInteger("iconWidth") & """ height=""" & CS.GetInteger("iconHeight") & """ sprites=""" & CS.GetInteger("iconSprites") & """ />"
                    result &= GetNodeBoolean(cp, "InIframe", CS.GetBoolean("InFrame"))
                    BlockEditTools = False
                    If CS.FieldOK("BlockEditTools") Then
                        BlockEditTools = CS.GetBoolean("BlockEditTools")
                    End If
                    result &= GetNodeBoolean(cp, "BlockEditTools", BlockEditTools)
                    '
                    ' Form XML
                    '
                    result &= GetNodeText(cp, "FormXML", CS.GetText("FormXML"))
                    '
                    NodeInnerText = ""
                    CS2.Open("Add-on Include Rules", "addonid=" & addonid)
                    Do While CS2.OK()
                        IncludedAddonID = CS2.GetInteger("IncludedAddonID")
                        CS3.Open("Add-ons", "ID=" & IncludedAddonID)
                        If CS3.OK() Then
                            Guid = CS3.GetText("ccGuid")
                            If Guid = "" Then
                                Guid = cp.Utils.CreateGuid()
                                Call CS3.SetField("ccGuid", Guid)
                            End If
                            result &= vbCrLf & vbTab & "<IncludeAddon name=""" & cp.Utils.EncodeHTML(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                        End If
                        Call CS3.Close()
                        Call CS2.GoNext()
                    Loop
                    Call CS2.Close()
                    '
                    result &= GetNodeBoolean(cp, "IsInline", CS.GetBoolean("IsInline"))
                    '
                    ' -- javascript (xmlnode may not match Db filename)
                    result &= GetNodeText(cp, "JavascriptInHead", CS.GetText("JSFilename"))
                    If (cp.Version > "4.2") Then
                        result &= GetNodeBoolean(cp, "javascriptForceHead", CS.GetBoolean("javascriptForceHead"))
                        result &= GetNodeText(cp, "JSHeadScriptSrc", CS.GetText("JSHeadScriptSrc"))
                    Else
                        result &= GetNodeBoolean(cp, "javascriptForceHead", False)
                        result &= GetNodeText(cp, "JSHeadScriptSrc", "")
                    End If
                    '
                    ' -- javascript deprecated
                    result &= GetNodeText(cp, "JSBodyScriptSrc", CS.GetText("JSBodyScriptSrc"), True)
                    result &= GetNodeText(cp, "JavascriptBodyEnd", CS.GetText("JavascriptBodyEnd"), True)
                    result &= GetNodeText(cp, "JavascriptOnLoad", CS.GetText("JavascriptOnLoad"), True)
                    '
                    ' -- Placements
                    result &= GetNodeBoolean(cp, "Content", CS.GetBoolean("Content"))
                    result &= GetNodeBoolean(cp, "Template", CS.GetBoolean("Template"))
                    result &= GetNodeBoolean(cp, "Email", CS.GetBoolean("Email"))
                    result &= GetNodeBoolean(cp, "Admin", CS.GetBoolean("Admin"))
                    result &= GetNodeBoolean(cp, "OnPageEndEvent", CS.GetBoolean("OnPageEndEvent"))
                    result &= GetNodeBoolean(cp, "OnPageStartEvent", CS.GetBoolean("OnPageStartEvent"))
                    result &= GetNodeBoolean(cp, "OnBodyStart", CS.GetBoolean("OnBodyStart"))
                    result &= GetNodeBoolean(cp, "OnBodyEnd", CS.GetBoolean("OnBodyEnd"))
                    result &= GetNodeBoolean(cp, "RemoteMethod", CS.GetBoolean("RemoteMethod"))
                    's = s & GetNodeBoolean( cp, "OnNewVisitEvent", CS.GetBoolean( "OnNewVisitEvent"))
                    '
                    ' -- Process
                    result &= GetNodeBoolean(cp, "ProcessRunOnce", processRunOnce)
                    result &= GetNodeInteger(cp, "ProcessInterval", CS.GetInteger("ProcessInterval"))
                    '
                    ' Meta
                    '
                    result &= GetNodeText(cp, "MetaDescription", CS.GetText("MetaDescription"))
                    result &= GetNodeText(cp, "OtherHeadTags", CS.GetText("OtherHeadTags"))
                    result &= GetNodeText(cp, "PageTitle", CS.GetText("PageTitle"))
                    result &= GetNodeText(cp, "RemoteAssetLink", CS.GetText("RemoteAssetLink"))
                    '
                    ' Styles
                    Styles = ""
                    If Not CS.GetBoolean("BlockDefaultStyles") Then
                        Styles = Trim(CS.GetText("StylesFilename"))
                    End If
                    StylesTest = Trim(CS.GetText("CustomStylesFilename"))
                    If StylesTest <> "" Then
                        If Styles <> "" Then
                            Styles = Styles & vbCrLf & StylesTest
                        Else
                            Styles = StylesTest
                        End If
                    End If
                    result &= GetNodeText(cp, "Styles", Styles)
                    result &= GetNodeText(cp, "styleslinkhref", CS.GetText("styleslinkhref"))
                    '
                    ' Scripting
                    '
                    NodeInnerText = Trim(CS.GetText("ScriptingCode"))
                    If NodeInnerText <> "" Then
                        NodeInnerText = vbCrLf & vbTab & vbTab & "<Code>" & EncodeCData(cp, NodeInnerText) & "</Code>"
                    End If
                    CS2.Open("Add-on Scripting Module Rules", "addonid=" & addonid)
                    Do While CS2.OK()
                        ScriptingModuleID = CS2.GetInteger("ScriptingModuleID")
                        CS3.Open("Scripting Modules", "ID=" & ScriptingModuleID)
                        If CS3.OK() Then
                            Guid = CS3.GetText("ccGuid")
                            If Guid = "" Then
                                Guid = cp.Utils.CreateGuid()
                                Call CS3.SetField("ccGuid", Guid)
                            End If
                            Return_IncludeModuleGuidList = Return_IncludeModuleGuidList & vbCrLf & Guid
                            NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<IncludeModule name=""" & cp.Utils.EncodeHTML(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                        End If
                        Call CS3.Close()
                        Call CS2.GoNext()
                    Loop
                    Call CS2.Close()
                    If NodeInnerText = "" Then
                        result &= vbCrLf & vbTab & "<Scripting Language=""" & CS.GetText("ScriptingLanguageID") & """ EntryPoint=""" & CS.GetText("ScriptingEntryPoint") & """ Timeout=""" & CS.GetText("ScriptingTimeout") & """/>"
                    Else
                        result &= vbCrLf & vbTab & "<Scripting Language=""" & CS.GetText("ScriptingLanguageID") & """ EntryPoint=""" & CS.GetText("ScriptingEntryPoint") & """ Timeout=""" & CS.GetText("ScriptingTimeout") & """>" & NodeInnerText & vbCrLf & vbTab & "</Scripting>"
                    End If
                    '
                    ' Shared Styles
                    '
                    CS2.Open("Shared Styles Add-on Rules", "addonid=" & addonid)
                    Do While CS2.OK()
                        styleId = CS2.GetInteger("styleId")
                        CS3.Open("shared styles", "ID=" & styleId)
                        If CS3.OK() Then
                            Guid = CS3.GetText("ccGuid")
                            If Guid = "" Then
                                Guid = cp.Utils.CreateGuid()
                                Call CS3.SetField("ccGuid", Guid)
                            End If
                            Return_IncludeSharedStyleGuidList = Return_IncludeSharedStyleGuidList & vbCrLf & Guid
                            result &= vbCrLf & vbTab & "<IncludeSharedStyle name=""" & cp.Utils.EncodeHTML(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                        End If
                        Call CS3.Close()
                        Call CS2.GoNext()
                    Loop
                    Call CS2.Close()
                    '
                    ' Process Triggers
                    '
                    NodeInnerText = ""
                    CS2.Open("Add-on Content Trigger Rules", "addonid=" & addonid)
                    Do While CS2.OK()
                        TriggerContentID = CS2.GetInteger("ContentID")
                        CS3.Open("content", "ID=" & TriggerContentID)
                        If CS3.OK() Then
                            Guid = CS3.GetText("ccGuid")
                            If Guid = "" Then
                                Guid = cp.Utils.CreateGuid()
                                Call CS3.SetField("ccGuid", Guid)
                            End If
                            NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<ContentChange name=""" & cp.Utils.EncodeHTML(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                        End If
                        Call CS3.Close()
                        Call CS2.GoNext()
                    Loop
                    Call CS2.Close()
                    If NodeInnerText <> "" Then
                        result &= vbCrLf & vbTab & "<ProcessTriggers>" & NodeInnerText & vbCrLf & vbTab & "</ProcessTriggers>"
                    End If
                    '
                    ' Editors
                    '
                    If cp.Content.IsField("Add-on Content Field Type Rules", "id") Then
                        NodeInnerText = ""
                        CS2.Open("Add-on Content Field Type Rules", "addonid=" & addonid)
                        Do While CS2.OK()
                            fieldTypeID = CS2.GetInteger("contentFieldTypeID")
                            fieldType = cp.Content.GetRecordName("Content Field Types", fieldTypeID)
                            If fieldType <> "" Then
                                NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<type>" & fieldType & "</type>"
                            End If
                            Call CS2.GoNext()
                        Loop
                        Call CS2.Close()
                        If NodeInnerText <> "" Then
                            result &= vbCrLf & vbTab & "<Editors>" & NodeInnerText & vbCrLf & vbTab & "</Editors>"
                        End If
                    End If
                    '
                    '
                    '
                    Guid = CS.GetText("ccGuid")
                    If Guid = "" Then
                        Guid = cp.Utils.CreateGuid()
                        Call CS.SetField("ccGuid", Guid)
                    End If
                    NavType = CS.GetText("NavTypeID")
                    If NavType = "" Then
                        NavType = "Add-on"
                    End If
                    result = "" _
                    & vbCrLf & vbTab & "<Addon name=""" & cp.Utils.EncodeHTML(addonName) & """ guid=""" & Guid & """ type=""" & NavType & """>" _
                    & tabIndent(cp, result) _
                    & vbCrLf & vbTab & "</Addon>"
                End If
                Call CS.Close()
            Catch ex As Exception
                errorReport(cp, ex, "GetAddonNode")
            End Try
            Return result
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple text node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="NodeContent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNodeText(cp As CPBaseClass, NodeName As String, NodeContent As String, Optional deprecated As Boolean = False) As String
            GetNodeText = ""
            Try
                Dim prefix As String = ""
                If (deprecated) Then
                    prefix = "<!-- deprecated -->"
                End If
                GetNodeText = ""
                If NodeContent = "" Then
                    GetNodeText = GetNodeText & vbCrLf & vbTab & prefix & "<" & NodeName & "></" & NodeName & ">"
                Else
                    GetNodeText = GetNodeText & vbCrLf & vbTab & prefix & "<" & NodeName & ">" & EncodeCData(cp, NodeContent) & "</" & NodeName & ">"
                End If
            Catch ex As Exception
                errorReport(cp, ex, "getNodeText")
            End Try
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple boolean node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="NodeContent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNodeBoolean(cp As CPBaseClass, NodeName As String, NodeContent As Boolean) As String
            GetNodeBoolean = ""
            Try
                GetNodeBoolean = vbCrLf & vbTab & "<" & NodeName & ">" & kmaGetYesNo(cp, NodeContent) & "</" & NodeName & ">"
            Catch ex As Exception
                errorReport(cp, ex, "GetNodeBoolean")
            End Try
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple integer node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="NodeContent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetNodeInteger(cp As CPBaseClass, NodeName As String, NodeContent As Integer) As String
            GetNodeInteger = ""
            Try
                GetNodeInteger = vbCrLf & vbTab & "<" & NodeName & ">" & CStr(NodeContent) & "</" & NodeName & ">"
            Catch ex As Exception
                errorReport(cp, ex, "GetNodeInteger")
            End Try
        End Function
        '
        '====================================================================================================
        Function replaceMany(cp As CPBaseClass, Source As String, ArrayOfSource() As String, ArrayOfReplacement() As String) As String
            replaceMany = ""
            Try
                Dim Count As Integer = UBound(ArrayOfSource) + 1
                replaceMany = Source
                For Pointer = 0 To Count - 1
                    replaceMany = Replace(replaceMany, ArrayOfSource(Pointer), ArrayOfReplacement(Pointer))
                Next
            Catch ex As Exception
                errorReport(cp, ex, "replaceMany")
            End Try
        End Function
        '
        '====================================================================================================
        Function encodeFilename(cp As CPBaseClass, Filename As String) As String
            encodeFilename = ""
            Try
                Dim Source() As String
                Dim Replacement() As String
                '
                Source = {"""", "*", "/", ":", "<", ">", "?", "\", "|"}
                Replacement = {"_", "_", "_", "_", "_", "_", "_", "_", "_"}
                '
                encodeFilename = replaceMany(cp, Filename, Source, Replacement)
                If Len(encodeFilename) > 254 Then
                    encodeFilename = Left(encodeFilename, 254)
                End If
            Catch ex As Exception
                errorReport(cp, ex, "encodeFilename")
            End Try
        End Function
        '
        '====================================================================================================
        Friend Sub GetLocalCollectionArgs(cp As CPBaseClass, CollectionGuid As String, ByRef Return_CollectionPath As String, ByRef Return_LastChagnedate As Date)
            Try
                Const CollectionListRootNode = "collectionlist"
                '
                Dim LocalPath As String
                Dim LocalGuid As String = ""
                Dim Doc As New Xml.XmlDocument
                Dim CollectionNode As Xml.XmlNode
                Dim LocalListNode As Xml.XmlNode
                Dim CollectionFound As Boolean
                Dim CollectionPath As String = ""
                Dim LastChangeDate As Date
                Dim MatchFound As Boolean
                Dim LocalName As String
                '
                MatchFound = False
                Return_CollectionPath = ""
                Return_LastChagnedate = Date.MinValue
                Call Doc.LoadXml(GetConfig(cp))
                If True Then
                    If LCase(Doc.DocumentElement.Name) <> LCase(CollectionListRootNode) Then
                        'Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, Hint=[" & Hint & "], The Collections.xml file has an invalid root node, [" & Doc.documentElement.name & "] was received and [" & CollectionListRootNode & "] was expected.")
                    Else
                        With Doc.DocumentElement
                            If LCase(.Name) <> "collectionlist" Then
                                'Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, basename was not collectionlist, [" & .name & "].")
                            Else
                                CollectionFound = False
                                'hint = hint & ",checking nodes [" & .childNodes.length & "]"
                                For Each LocalListNode In .ChildNodes
                                    LocalName = "no name found"
                                    LocalPath = ""
                                    Select Case LCase(LocalListNode.Name)
                                        Case "collection"
                                            LocalGuid = ""
                                            For Each CollectionNode In LocalListNode.ChildNodes
                                                Select Case LCase(CollectionNode.Name)
                                                    Case "name"
                                                        '
                                                        LocalName = LCase(CollectionNode.InnerText)
                                                    Case "guid"
                                                        '
                                                        LocalGuid = LCase(CollectionNode.InnerText)
                                                    Case "path"
                                                        '
                                                        CollectionPath = LCase(CollectionNode.InnerText)
                                                    Case "lastchangedate"
                                                        LastChangeDate = cp.Utils.EncodeDate(CollectionNode.InnerText)
                                                End Select
                                            Next
                                    End Select
                                    'hint = hint & ",checking node [" & LocalName & "]"
                                    If LCase(CollectionGuid) = LocalGuid Then
                                        Return_CollectionPath = CollectionPath
                                        Return_LastChagnedate = LastChangeDate
                                        'Call AppendClassLogFile("Server", "GetCollectionConfigArg", "GetLocalCollectionArgs, match found, CollectionName=" & LocalName & ", CollectionPath=" & CollectionPath & ", LastChangeDate=" & LastChangeDate)
                                        MatchFound = True
                                        Exit For
                                    End If
                                Next
                            End If
                        End With
                    End If
                End If
                If Not MatchFound Then
                    'Call AppendClassLogFile("Server", "GetCollectionConfigArg", "GetLocalCollectionArgs, no local collection match found, Hint=[" & Hint & "]")
                End If
            Catch ex As Exception
                errorReport(cp, ex, "GetLocalCollectionArgs")
            End Try
        End Sub
        '
        '====================================================================================================
        Public Function GetConfig(cp As CPBaseClass) As String
            GetConfig = ""
            Try
                Dim AddonPath As String
                '
                AddonPath = cp.Site.PhysicalInstallPath & "\addons"
                AddonPath = AddonPath & "\Collections.xml"
                GetConfig = cp.File.Read(AddonPath)
            Catch ex As Exception
                errorReport(cp, ex, "GetConfig")
            End Try
        End Function
        '
        '====================================================================================================
        Private Function AddCompatibilityResources(cp As CPBaseClass, CollectionPath As String, ArchiveFilename As String, SubPath As String) As String
            AddCompatibilityResources = ""
            Dim s As String = ""
            Try
                Dim AddFilename As String
                Dim FileExt As String
                Dim FileList As String
                Dim Files() As String
                Dim Filename As String
                Dim Ptr As Integer
                Dim FileArgs() As String
                Dim FolderList As String
                Dim Folders() As String
                Dim FolderArgs() As String
                Dim Folder As String
                Dim Pos As Integer
                '
                ' Process all SubPaths
                '
                FolderList = cp.File.folderList(CollectionPath & SubPath)
                If FolderList <> "" Then
                    Folders = Split(FolderList, vbCrLf)
                    For Ptr = 0 To UBound(Folders)
                        Folder = Folders(Ptr)
                        If Folder <> "" Then
                            FolderArgs = Split(Folders(Ptr), ",")
                            Folder = FolderArgs(0)
                            If Folder <> "" Then
                                s = s & AddCompatibilityResources(cp, CollectionPath, ArchiveFilename, SubPath & Folder & "\")
                            End If
                        End If
                    Next
                End If
                '
                ' Process files in this path
                '
                'Set Remote = CreateObject("ccRemote.RemoteClass")
                FileList = cp.File.fileList(CollectionPath)
                If FileList <> "" Then
                    Files = Split(FileList, vbCrLf)
                    For Ptr = 0 To UBound(Files)
                        Filename = Files(Ptr)
                        If Filename <> "" Then
                            FileArgs = Split(Filename, ",")
                            If UBound(FileArgs) > 0 Then
                                Filename = FileArgs(0)
                                Pos = InStrRev(Filename, ".")
                                FileExt = ""
                                If Pos > 0 Then
                                    FileExt = Mid(Filename, Pos + 1)
                                End If
                                If LCase(Filename) = "collection.hlp" Then
                                    '
                                    ' legacy help system, ignore this file
                                    '
                                ElseIf LCase(FileExt) = "xml" Then
                                    '
                                    ' compatibility resources can not include an xml file in the wwwroot
                                    '
                                ElseIf InStr(1, CollectionPath, "\ContensiveFiles\", vbTextCompare) <> 0 Then
                                    '
                                    ' Content resources
                                    '
                                    s = s & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""content"" path=""" & cp.Utils.EncodeHTML(SubPath) & """ />"
                                    AddFilename = CollectionPath & SubPath & "\" & Filename
                                    'Call zipFile(ArchiveFilename, AddFilename)
                                    'Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                    'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                ElseIf LCase(FileExt) = "dll" Then
                                    '
                                    ' Executable resources
                                    '
                                    s = s & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""executable"" path=""" & cp.Utils.EncodeHTML(SubPath) & """ />"
                                    AddFilename = CollectionPath & SubPath & "\" & Filename
                                    'Call zipFile(ArchiveFilename, AddFilename)
                                    'Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                    'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                Else
                                    '
                                    ' www resources
                                    '
                                    s = s & vbCrLf & vbTab & "<Resource name=""" & cp.Utils.EncodeHTML(Filename) & """ type=""www"" path=""" & cp.Utils.EncodeHTML(SubPath) & """ />"
                                    AddFilename = CollectionPath & SubPath & "\" & Filename
                                    'Call zipFile(ArchiveFilename, AddFilename)
                                    'Call runAtServer("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                    'Call Remote.executeCmd("zipfile", "archive=" & kmaEncodeRequestVariable(ArchiveFilename) & "&add=" & kmaEncodeRequestVariable(AddFilename))
                                End If
                            End If
                        End If
                    Next
                End If
                '
                AddCompatibilityResources = s
            Catch ex As Exception
                errorReport(cp, ex, "GetNodeInteger")
            End Try
        End Function
        '
        '====================================================================================================
        Friend Function EncodeCData(cp As CPBaseClass, Source As String) As String
            EncodeCData = ""
            Try
                EncodeCData = Source
                If EncodeCData <> "" Then
                    EncodeCData = "<![CDATA[" & Replace(EncodeCData, "]]>", "]]]]><![CDATA[>") & "]]>"
                End If
            Catch ex As Exception
                errorReport(cp, ex, "EncodeCData")
            End Try
        End Function
        '
        '====================================================================================================
        Public Function kmaGetYesNo(cp As CPBaseClass, Key As Boolean) As String
            If Key Then
                kmaGetYesNo = "Yes"
            Else
                kmaGetYesNo = "No"
            End If
        End Function
        '
        '=======================================================================================
        ''' <summary>
        ''' zip
        ''' </summary>
        ''' <param name="PathFilename"></param>
        ''' <remarks></remarks>
        Public Sub UnzipFile(cp As CPBaseClass, ByVal PathFilename As String)
            Try
                '
                Dim fastZip As ICSharpCode.SharpZipLib.Zip.FastZip = New ICSharpCode.SharpZipLib.Zip.FastZip()
                Dim fileFilter As String = Nothing

                fastZip.ExtractZip(PathFilename, getPath(cp, PathFilename), fileFilter)                '
            Catch ex As Exception
                errorReport(cp, ex, "UnzipFile")
            End Try
        End Sub        '
        '
        '=======================================================================================
        ''' <summary>
        ''' unzip
        ''' </summary>
        ''' <param name="archivePathFilename"></param>
        ''' <param name="addPathFilename"></param>
        ''' <remarks></remarks>
        Public Sub zipFile(cp As CPBaseClass, archivePathFilename As String, ByVal addPathFilename As List(Of String))
            Try
                '
                'Dim fastZip As FastZip = New ICSharpCode.SharpZipLib.Zip.FastZip()
                'Dim fileFilter As String = Nothing
                'Dim recurse As Boolean = True
                'Dim archivepath As String = getPath(archivePathFilename)
                'Dim archiveFilename As String = GetFilename(archivePathFilename)
                '
                Dim z As Zip.ZipFile
                If cp.File.fileExists(archivePathFilename) Then
                    '
                    ' update existing zip with list of files
                    '
                    z = New Zip.ZipFile(archivePathFilename)
                Else
                    '
                    ' create new zip
                    '
                    z = Zip.ZipFile.Create(archivePathFilename)
                End If
                z.BeginUpdate()
                For Each pathFilename In addPathFilename
                    z.Add(pathFilename, System.IO.Path.GetFileName(pathFilename))
                Next
                z.CommitUpdate()
                z.Close()
                'fastZip.CreateZip(archivePathFilename, addPathFilename, recurse, fileFilter)
            Catch ex As Exception
                errorReport(cp, ex, "zipFile")
            End Try
        End Sub        '
        '
        '=======================================================================================
        Private Function getPath(cp As CPBaseClass, ByVal pathFilename As String) As String
            getPath = ""
            Try
                Dim Position As Integer
                '
                Position = InStrRev(pathFilename, "\")
                If Position <> 0 Then
                    getPath = Mid(pathFilename, 1, Position)
                End If
            Catch ex As Exception
                errorReport(cp, ex, "getPath")
            End Try
        End Function
        '
        '=======================================================================================
        Public Function GetFilename(cp As CPBaseClass, ByVal PathFilename As String) As String
            Dim Position As Integer
            '
            GetFilename = PathFilename
            Position = InStrRev(GetFilename, "/")
            If Position <> 0 Then
                GetFilename = Mid(GetFilename, Position + 1)
            End If
        End Function
        '
        '=======================================================================================
        '
        '   Indent every line by 1 tab
        '
        Public Function tabIndent(cp As CPBaseClass, Source As String) As String
            Dim posStart As Integer
            Dim posEnd As Integer
            Dim pre As String
            Dim post As String
            Dim target As String
            '
            posStart = InStr(1, Source, "<![CDATA[", CompareMethod.Text)
            If posStart = 0 Then
                '
                ' no cdata
                '
                posStart = InStr(1, Source, "<textarea", CompareMethod.Text)
                If posStart = 0 Then
                    '
                    ' no textarea
                    '
                    tabIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
                Else
                    '
                    ' text area found, isolate it and indent before and after
                    '
                    posEnd = InStr(posStart, Source, "</textarea>", CompareMethod.Text)
                    pre = Mid(Source, 1, posStart - 1)
                    If posEnd = 0 Then
                        target = Mid(Source, posStart)
                        post = ""
                    Else
                        target = Mid(Source, posStart, posEnd - posStart + Len("</textarea>"))
                        post = Mid(Source, posEnd + Len("</textarea>"))
                    End If
                    tabIndent = tabIndent(cp, pre) & target & tabIndent(cp, post)
                End If
            Else
                '
                ' cdata found, isolate it and indent before and after
                '
                posEnd = InStr(posStart, Source, "]]>", CompareMethod.Text)
                pre = Mid(Source, 1, posStart - 1)
                If posEnd = 0 Then
                    target = Mid(Source, posStart)
                    post = ""
                Else
                    target = Mid(Source, posStart, posEnd - posStart + Len("]]>"))
                    post = Mid(Source, posEnd + 3)
                End If
                tabIndent = tabIndent(cp, pre) & target & tabIndent(cp, post)
            End If
            '    kmaIndent = Source
            '    If InStr(1, kmaIndent, "<textarea", vbTextCompare) = 0 Then
            '        kmaIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
            '    End If
        End Function

    End Class
End Namespace
