
Imports Contensive.Addons.AddonManager51.Models
Imports Contensive.BaseClasses
Imports ICSharpCode.SharpZipLib

Namespace Contensive.Addons.AddonManager51
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
                                    Dim CollectionFilename As String = createCollectionZip_returnCdnPathFilename(CP, CollectionID)
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
                            form.rowValue = CP.Html.SelectContent(RequestNameCollectionID, "0", "Add-on Collections", "", "Select Collection To Export", "form-control")
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
        ''' <summary>
        ''' create the colleciton zip file and return the pathFilename in the Cdn
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="CollectionID"></param>
        ''' <returns></returns>
        Private Function createCollectionZip_returnCdnPathFilename(cp As CPBaseClass, CollectionID As Integer) As String
            Dim cdnExportZip_Filename As String = ""
            Try
                Dim CS As CPCSBaseClass = cp.CSNew()
                CS.OpenRecord("Add-on Collections", CollectionID)
                If Not CS.OK() Then
                    Call cp.UserError.Add("The collection you selected could not be found")
                Else
                    Dim collectionXml As String = "<?xml version=""1.0"" encoding=""windows-1252""?>"
                    '
                    Dim CollectionName As String = CS.GetText("name")
                    Dim CollectionGuid As String = CS.GetText("ccGuid")
                    If CollectionGuid = "" Then
                        CollectionGuid = cp.Utils.CreateGuid()
                        Call CS.SetField("ccGuid", CollectionGuid)
                    End If
                    Dim onInstallAddonGuid As String = ""
                    If (CS.FieldOK("onInstallAddonId")) Then
                        Dim onInstallAddonId As Integer = CS.GetInteger("onInstallAddonId")
                        If (onInstallAddonId > 0) Then
                            Dim addon As AddonModel = AddonModel.create(cp, onInstallAddonId)
                            onInstallAddonGuid = addon.ccguid
                        End If
                    End If
                    collectionXml &= vbCrLf & "<Collection"
                    collectionXml &= " name=""" & CollectionName & """"
                    collectionXml &= " guid=""" & CollectionGuid & """"
                    collectionXml &= " system=""" & getYesNo(cp, CS.GetBoolean("system")) & """"
                    collectionXml &= " updatable=""" & getYesNo(cp, CS.GetBoolean("updatable")) & """"
                    collectionXml &= " blockNavigatorNode=""" & getYesNo(cp, CS.GetBoolean("blockNavigatorNode")) & """"
                    collectionXml &= " onInstallAddonGuid=""" & onInstallAddonGuid & """"
                    collectionXml &= ">"
                    '
                    ' Archive Filenames
                    '   copy all files to be included into the cdnExportFilesPath folder
                    '   build the tmp zip file
                    '   copy it to the cdnZip file
                    '
                    Dim tempExportPath As String = "CollectionExport" & Guid.NewGuid().ToString() & "\"
                    Dim tempExportXml_Filename As String = encodeFilename(cp, CollectionName & ".xml")
                    Dim tempExportZip_Filename As String = encodeFilename(cp, CollectionName & ".zip")
                    cdnExportZip_Filename = encodeFilename(cp, CollectionName & ".zip")
                    '
                    ' Delete old archive file
                    'cp.TempFiles.DeleteFile(tempExportXml_Filename)
                    'cp.TempFiles.DeleteFile(tempExportZipPathFilename)
                    'cp.CdnFiles.DeleteFile(cdnExportZipPathFilename)
                    '
                    '
                    ' Build executable file list Resource Node so executables can be added to addons for Version40compatibility
                    '   but save it for the end, executableFileList
                    '
                    'Call Main.testpoint("getCollection, 400")
                    Dim AddonPath As String = "addons\"
                    Dim FileList As String = CS.GetText("execFileList")
                    Dim Path As String
                    Dim Filename As String
                    Dim PathFilename As String
                    Dim Ptr As Integer
                    Dim Files() As String
                    Dim ResourceCnt As Integer
                    'Dim ContentName As String
                    Dim Pos As Integer
                    Dim tempPathFileList As New List(Of String)
                    'Dim PhysicalWWWPath As String
                    Dim CollectionPath As String = ""
                    Dim ExecFileListNode As String = ""
                    If FileList <> "" Then
                        Dim LastChangeDate As Date
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
                                Dim ManualFilename As String = ""
                                If LCase(Filename) <> LCase(ManualFilename) Then
                                    'AddFilename = AddonPath & CollectionPath & Filename
                                    cp.PrivateFiles.Copy(AddonPath & CollectionPath & Filename, tempExportPath & Filename, cp.TempFiles)
                                    If Not tempPathFileList.Contains(tempExportPath & Filename) Then
                                        tempPathFileList.Add(tempExportPath & Filename)
                                        ExecFileListNode = ExecFileListNode & vbCrLf & vbTab & "<Resource name=""" & System.Net.WebUtility.HtmlEncode(Filename) & """ type=""executable"" path=""" & System.Net.WebUtility.HtmlEncode(Path) & """ />"
                                    End If
                                End If
                                ResourceCnt = ResourceCnt + 1
                            End If
                        Next
                    End If
                    'Call Main.testpoint("getCollection, 500")
                    'If (ResourceCnt = 0) And (CollectionPath <> "") Then
                    '    '
                    '    ' If no resources were in the collection record, this might be an old installation
                    '    ' Add all .dll files in the CollectionPath
                    '    '
                    '    ExecFileListNode = ExecFileListNode & AddCompatibilityResources(cp, AddonPath & CollectionPath, cdnTempZipPathFilename, "")
                    'End If
                    '
                    ' helpLink
                    '
                    If CS.FieldOK("HelpLink") Then
                        collectionXml = collectionXml & vbCrLf & vbTab & "<HelpLink>" & System.Net.WebUtility.HtmlEncode(CS.GetText("HelpLink")) & "</HelpLink>"
                    End If
                    '
                    ' Help
                    '
                    collectionXml = collectionXml & vbCrLf & vbTab & "<Help>" & System.Net.WebUtility.HtmlEncode(CS.GetText("Help")) & "</Help>"
                    Dim CS2 As CPCSBaseClass = cp.CSNew()
                    '
                    ' Addons
                    '
                    CS2.Open("Add-ons", "collectionid=" & CollectionID, "name", True, "id")
                    Dim IncludeModuleGuidList As String = ""
                    Dim IncludeSharedStyleGuidList As String = ""
                    Do While CS2.OK()
                        collectionXml = collectionXml & getAddonNode(cp, CS2.GetInteger("id"), IncludeModuleGuidList, IncludeSharedStyleGuidList)
                        Call CS2.GoNext()
                    Loop
                    '
                    ' Data Records
                    '
                    Dim DataRecordList As String = CS.GetText("DataRecordList")
                    If DataRecordList <> "" Then
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
                                        Dim CSData As CPCSBaseClass = cp.CSNew()
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
                                            Dim csFields As CPCSBaseClass = cp.CSNew()
                                            Dim fieldLookupListValue As String = ""
                                            Dim fieldNames() As String = {}
                                            Dim fieldTypes() As Integer = {}
                                            Dim fieldLookupContent() As String = {}
                                            Dim fieldLookupList() As String = {}
                                            Dim FieldLookupContentName As String
                                            Dim FieldTypeNumber As Integer
                                            Dim FieldName As String
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
                                                Dim fieldPtr As Integer
                                                For fieldPtr = 0 To fieldCnt - 1
                                                    FieldName = fieldNames(fieldPtr)
                                                    FieldTypeNumber = cp.Utils.EncodeInteger(fieldTypes(fieldPtr))
                                                    'Dim ContentID As Integer
                                                    Dim FieldValue As String
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
                                                            Dim FieldValueInteger As Integer = CSData.GetInteger(FieldName)
                                                            If (FieldValueInteger <> 0) Then
                                                                FieldLookupContentName = fieldLookupContent(fieldPtr)
                                                                fieldLookupListValue = fieldLookupList(fieldPtr)
                                                                If (FieldLookupContentName <> "") Then
                                                                    '
                                                                    ' content lookup
                                                                    '
                                                                    If cp.Content.IsField(FieldLookupContentName, "ccguid") Then
                                                                        Dim CSlookup As CPCSBaseClass = cp.CSNew()
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
                                                    FieldNodes = FieldNodes & vbCrLf & vbTab & "<field name=""" & System.Net.WebUtility.HtmlEncode(FieldName) & """>" & FieldValue & "</field>"
                                                Next
                                                RecordNodes = "" _
                                                    & RecordNodes _
                                                    & vbCrLf & vbTab & "<record content=""" & System.Net.WebUtility.HtmlEncode(DataContentName) & """ guid=""" & DataRecordGuid & """ name=""" & System.Net.WebUtility.HtmlEncode(DataRecordName) & """>" _
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
                    Dim Node As String
                    '
                    ' CDef
                    '
                    'Call Main.testpoint("getCollection, 700")
                    For Each content As Models.ContentModel In Models.ContentModel.createListFromCollection(cp, CollectionID)
                        Dim reload As Boolean = False
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
                        Dim Modules() As String = Split(IncludeModuleGuidList, vbCrLf)
                        For Ptr = 0 To UBound(Modules)
                            Dim ModuleGuid As String = Modules(Ptr)
                            If ModuleGuid <> "" Then
                                CS2.Open("Scripting Modules", "ccguid=" & cp.Db.EncodeSQLText(ModuleGuid))
                                If CS2.OK() Then
                                    Dim Code As String = Trim(CS2.GetText("code"))
                                    Code = EncodeCData(cp, Code)
                                    collectionXml = collectionXml & vbCrLf & vbTab & "<ScriptingModule Name=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("name")) & """ guid=""" & ModuleGuid & """>" & Code & "</ScriptingModule>"
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
                                        & " Name=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("name")) & """" _
                                        & " guid=""" & recordGuid & """" _
                                        & " alwaysInclude=""" & CS2.GetBoolean("alwaysInclude") & """" _
                                        & " prefix=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("prefix")) & """" _
                                        & " suffix=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("suffix")) & """" _
                                        & " sortOrder=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("sortOrder")) & """" _
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
                    Dim CS3 As CPCSBaseClass = cp.CSNew()
                    If CS3.Open("Add-on Collection Parent Rules", "parentid=" & CollectionID) Then
                        Do
                            CS2.OpenRecord("Add-on Collections", CS3.GetInteger("childid"))
                            If CS2.OK() Then
                                Dim Guid As String = CS2.GetText("ccGuid")
                                If Guid = "" Then
                                    Guid = cp.Utils.CreateGuid()
                                    Call CS2.SetField("ccGuid", Guid)
                                End If
                                Node = Node & vbCrLf & vbTab & "<ImportCollection name=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("name")) & """>" & Guid & "</ImportCollection>"
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
                                    If tempPathFileList.Contains(tempExportPath & Filename) Then
                                        Call cp.UserError.Add("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                                    Else
                                        cp.WwwFiles.Copy(PathFilename, tempExportPath & Filename, cp.TempFiles)
                                        tempPathFileList.Add(tempExportPath & Filename)
                                        collectionXml = collectionXml & vbCrLf & vbTab & "<Resource name=""" & System.Net.WebUtility.HtmlEncode(Filename) & """ type=""www"" path=""" & System.Net.WebUtility.HtmlEncode(Path) & """ />"
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
                                'PathFilename = Replace(PathFilename, "/", "\")
                                'If Left(PathFilename, 1) = "\" Then
                                '    PathFilename = Mid(PathFilename, 2)
                                'End If
                                If tempPathFileList.Contains(tempExportPath & Filename) Then
                                    Call cp.UserError.Add("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                                Else
                                    cp.CdnFiles.Copy(PathFilename, tempExportPath + Filename, cp.TempFiles)
                                    tempPathFileList.Add(tempExportPath & Filename)
                                    collectionXml = collectionXml & vbCrLf & vbTab & "<Resource name=""" & System.Net.WebUtility.HtmlEncode(Filename) & """ type=""content"" path=""" & System.Net.WebUtility.HtmlEncode(Path) & """ />"
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
                    Call cp.TempFiles.Save(tempExportPath & tempExportXml_Filename, collectionXml)
                    If Not tempPathFileList.Contains(tempExportPath & tempExportXml_Filename) Then
                        tempPathFileList.Add(tempExportPath & tempExportXml_Filename)
                    End If
                    '
                    ' -- zip up the folder to make the collection zip file in temp filesystem
                    Call zipTempCdnFile(cp, tempExportPath & tempExportZip_Filename, tempPathFileList)
                    '
                    ' -- copy the collection zip file to the cdn filesystem as the download link
                    cp.TempFiles.Copy(tempExportPath & tempExportZip_Filename, cdnExportZip_Filename, cp.CdnFiles)
                    '
                    ' -- delete the temp folder
                    cp.TempFiles.DeleteFolder(tempExportPath)
                End If
            Catch ex As Exception
                errorReport(cp, ex, "GetCollection")
            End Try
            Return cdnExportZip_Filename
        End Function
        '
        '====================================================================================================

        Private Function getAddonNode(cp As CPBaseClass, addonid As Integer, ByRef Return_IncludeModuleGuidList As String, ByRef Return_IncludeSharedStyleGuidList As String) As String
            Dim result As String = ""
            Try
                Using CS As CPCSBaseClass = cp.CSNew()
                    If CS.OpenRecord("Add-ons", addonid) Then
                        Dim addonName As String = CS.GetText("name")
                        Dim processRunOnce As Boolean = CS.GetBoolean("ProcessRunOnce")
                        If ((LCase(addonName) = "oninstall") Or (LCase(addonName) = "_oninstall")) Then
                            processRunOnce = True
                        End If
                        '
                        ' ActiveX DLL node is being deprecated. This should be in the collection resource section
                        result &= getNodeText(cp, "Copy", CS.GetText("Copy"))
                        result &= getNodeText(cp, "CopyText", CS.GetText("CopyText"))
                        '
                        ' DLL
                        result &= getNodeText(cp, "ActiveXProgramID", CS.GetText("objectprogramid"), True)
                        result &= getNodeText(cp, "DotNetClass", CS.GetText("DotNetClass"))
                        '
                        ' Features
                        result &= getNodeText(cp, "ArgumentList", CS.GetText("ArgumentList"))
                        result &= getNodeBoolean(cp, "AsAjax", CS.GetBoolean("AsAjax"))
                        result &= getNodeBoolean(cp, "Filter", CS.GetBoolean("Filter"))
                        result &= getNodeText(cp, "Help", CS.GetText("Help"))
                        result &= getNodeText(cp, "HelpLink", CS.GetText("HelpLink"))
                        result &= vbCrLf & vbTab & "<Icon Link=""" & CS.GetText("iconfilename") & """ width=""" & CS.GetInteger("iconWidth") & """ height=""" & CS.GetInteger("iconHeight") & """ sprites=""" & CS.GetInteger("iconSprites") & """ />"
                        result &= getNodeBoolean(cp, "InIframe", CS.GetBoolean("InFrame"))
                        result &= If(CS.FieldOK("BlockEditTools"), getNodeBoolean(cp, "BlockEditTools", CS.GetBoolean("BlockEditTools")), "")
                        result &= If(CS.FieldOK("aliasList"), getNodeText(cp, "AliasList", CS.GetText("aliasList")), "")
                        '
                        ' -- Form XML
                        result &= getNodeText(cp, "FormXML", CS.GetText("FormXML"))
                        '
                        ' -- addon dependencies
                        Using CS2 As CPCSBaseClass = cp.CSNew()
                            CS2.Open("Add-on Include Rules", "addonid=" & addonid)
                            Do While CS2.OK()
                                Dim IncludedAddonID As Integer = CS2.GetInteger("IncludedAddonID")
                                Using CS3 As CPCSBaseClass = cp.CSNew()
                                    CS3.Open("Add-ons", "ID=" & IncludedAddonID)
                                    If CS3.OK() Then
                                        Dim Guid As String = CS3.GetText("ccGuid")
                                        If Guid = "" Then
                                            Guid = cp.Utils.CreateGuid()
                                            Call CS3.SetField("ccGuid", Guid)
                                        End If
                                        result &= vbCrLf & vbTab & "<IncludeAddon name=""" & System.Net.WebUtility.HtmlEncode(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                                    End If
                                    Call CS3.Close()
                                End Using
                                Call CS2.GoNext()
                            Loop
                            Call CS2.Close()
                        End Using
                        '
                        ' -- is inline/block
                        result &= getNodeBoolean(cp, "IsInline", CS.GetBoolean("IsInline"))
                        '
                        ' -- javascript (xmlnode may not match Db filename)
                        result &= getNodeText(cp, "JavascriptInHead", CS.GetText("JSFilename"))
                        result &= getNodeBoolean(cp, "javascriptForceHead", CS.GetBoolean("javascriptForceHead"))
                        result &= getNodeText(cp, "JSHeadScriptSrc", CS.GetText("JSHeadScriptSrc"))
                        '
                        ' -- javascript deprecated
                        result &= getNodeText(cp, "JSBodyScriptSrc", CS.GetText("JSBodyScriptSrc"), True)
                        result &= getNodeText(cp, "JavascriptBodyEnd", CS.GetText("JavascriptBodyEnd"), True)
                        result &= getNodeText(cp, "JavascriptOnLoad", CS.GetText("JavascriptOnLoad"), True)
                        '
                        ' -- Placements
                        result &= getNodeBoolean(cp, "Content", CS.GetBoolean("Content"))
                        result &= getNodeBoolean(cp, "Template", CS.GetBoolean("Template"))
                        result &= getNodeBoolean(cp, "Email", CS.GetBoolean("Email"))
                        result &= getNodeBoolean(cp, "Admin", CS.GetBoolean("Admin"))
                        result &= getNodeBoolean(cp, "OnPageEndEvent", CS.GetBoolean("OnPageEndEvent"))
                        result &= getNodeBoolean(cp, "OnPageStartEvent", CS.GetBoolean("OnPageStartEvent"))
                        result &= getNodeBoolean(cp, "OnBodyStart", CS.GetBoolean("OnBodyStart"))
                        result &= getNodeBoolean(cp, "OnBodyEnd", CS.GetBoolean("OnBodyEnd"))
                        result &= getNodeBoolean(cp, "RemoteMethod", CS.GetBoolean("RemoteMethod"))
                        result &= If(CS.FieldOK("Diagnostic"), getNodeBoolean(cp, "Diagnostic", CS.GetBoolean("Diagnostic")), "")
                        '
                        ' -- Process
                        result &= getNodeBoolean(cp, "ProcessRunOnce", processRunOnce)
                        result &= GetNodeInteger(cp, "ProcessInterval", CS.GetInteger("ProcessInterval"))
                        '
                        ' Meta
                        '
                        result &= getNodeText(cp, "MetaDescription", CS.GetText("MetaDescription"))
                        result &= getNodeText(cp, "OtherHeadTags", CS.GetText("OtherHeadTags"))
                        result &= getNodeText(cp, "PageTitle", CS.GetText("PageTitle"))
                        result &= getNodeText(cp, "RemoteAssetLink", CS.GetText("RemoteAssetLink"))
                        '
                        ' Styles
                        Dim Styles As String = ""
                        If Not CS.GetBoolean("BlockDefaultStyles") Then
                            Styles = Trim(CS.GetText("StylesFilename"))
                        End If
                        Dim StylesTest As String = Trim(CS.GetText("CustomStylesFilename"))
                        If StylesTest <> "" Then
                            If Styles <> "" Then
                                Styles = Styles & vbCrLf & StylesTest
                            Else
                                Styles = StylesTest
                            End If
                        End If
                        result &= getNodeText(cp, "Styles", Styles)
                        result &= getNodeText(cp, "styleslinkhref", CS.GetText("styleslinkhref"))
                        '
                        '
                        ' Scripting
                        '
                        Dim NodeInnerText As String = Trim(CS.GetText("ScriptingCode"))
                        If NodeInnerText <> "" Then
                            NodeInnerText = vbCrLf & vbTab & vbTab & "<Code>" & EncodeCData(cp, NodeInnerText) & "</Code>"
                        End If
                        Using CS2 As CPCSBaseClass = cp.CSNew()
                            CS2.Open("Add-on Scripting Module Rules", "addonid=" & addonid)
                            Do While CS2.OK()
                                Dim ScriptingModuleID As Integer = CS2.GetInteger("ScriptingModuleID")
                                Using CS3 As CPCSBaseClass = cp.CSNew()
                                    CS3.Open("Scripting Modules", "ID=" & ScriptingModuleID)
                                    If CS3.OK() Then
                                        Dim Guid As String = CS3.GetText("ccGuid")
                                        If Guid = "" Then
                                            Guid = cp.Utils.CreateGuid()
                                            Call CS3.SetField("ccGuid", Guid)
                                        End If
                                        Return_IncludeModuleGuidList = Return_IncludeModuleGuidList & vbCrLf & Guid
                                        NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<IncludeModule name=""" & System.Net.WebUtility.HtmlEncode(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                                    End If
                                    Call CS3.Close()
                                End Using
                                Call CS2.GoNext()
                            Loop
                            Call CS2.Close()
                        End Using
                        If NodeInnerText = "" Then
                            result &= vbCrLf & vbTab & "<Scripting Language=""" & CS.GetText("ScriptingLanguageID") & """ EntryPoint=""" & CS.GetText("ScriptingEntryPoint") & """ Timeout=""" & CS.GetText("ScriptingTimeout") & """/>"
                        Else
                            result &= vbCrLf & vbTab & "<Scripting Language=""" & CS.GetText("ScriptingLanguageID") & """ EntryPoint=""" & CS.GetText("ScriptingEntryPoint") & """ Timeout=""" & CS.GetText("ScriptingTimeout") & """>" & NodeInnerText & vbCrLf & vbTab & "</Scripting>"
                        End If
                        '
                        ' Shared Styles
                        '
                        Using CS2 As CPCSBaseClass = cp.CSNew()
                            CS2.Open("Shared Styles Add-on Rules", "addonid=" & addonid)
                            Do While CS2.OK()
                                Dim styleId As Integer = CS2.GetInteger("styleId")
                                Using CS3 As CPCSBaseClass = cp.CSNew()
                                    CS3.Open("shared styles", "ID=" & styleId)
                                    If CS3.OK() Then
                                        Dim Guid As String = CS3.GetText("ccGuid")
                                        If Guid = "" Then
                                            Guid = cp.Utils.CreateGuid()
                                            Call CS3.SetField("ccGuid", Guid)
                                        End If
                                        Return_IncludeSharedStyleGuidList = Return_IncludeSharedStyleGuidList & vbCrLf & Guid
                                        result &= vbCrLf & vbTab & "<IncludeSharedStyle name=""" & System.Net.WebUtility.HtmlEncode(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                                    End If
                                    Call CS3.Close()
                                End Using
                                Call CS2.GoNext()
                            Loop
                            Call CS2.Close()
                        End Using
                        '
                        ' Process Triggers
                        '
                        NodeInnerText = ""
                        Using CS2 As CPCSBaseClass = cp.CSNew()
                            CS2.Open("Add-on Content Trigger Rules", "addonid=" & addonid)
                            Do While CS2.OK()
                                Dim TriggerContentID As Integer = CS2.GetInteger("ContentID")
                                Using CS3 As CPCSBaseClass = cp.CSNew()
                                    CS3.Open("content", "ID=" & TriggerContentID)
                                    If CS3.OK() Then
                                        Dim Guid As String = CS3.GetText("ccGuid")
                                        If Guid = "" Then
                                            Guid = cp.Utils.CreateGuid()
                                            Call CS3.SetField("ccGuid", Guid)
                                        End If
                                        NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<ContentChange name=""" & System.Net.WebUtility.HtmlEncode(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                                    End If
                                    Call CS3.Close()
                                End Using
                                Call CS2.GoNext()
                            Loop
                            Call CS2.Close()
                        End Using
                        If NodeInnerText <> "" Then
                            result &= vbCrLf & vbTab & "<ProcessTriggers>" & NodeInnerText & vbCrLf & vbTab & "</ProcessTriggers>"
                        End If
                        '
                        ' Editors
                        '
                        If cp.Content.IsField("Add-on Content Field Type Rules", "id") Then
                            NodeInnerText = ""
                            Using CS2 As CPCSBaseClass = cp.CSNew()
                                CS2.Open("Add-on Content Field Type Rules", "addonid=" & addonid)
                                Do While CS2.OK()
                                    Dim fieldTypeID As Integer = CS2.GetInteger("contentFieldTypeID")
                                    Dim fieldType As String = cp.Content.GetRecordName("Content Field Types", fieldTypeID)
                                    If fieldType <> "" Then
                                        NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<type>" & fieldType & "</type>"
                                    End If
                                    Call CS2.GoNext()
                                Loop
                                Call CS2.Close()
                            End Using
                            If NodeInnerText <> "" Then
                                result &= vbCrLf & vbTab & "<Editors>" & NodeInnerText & vbCrLf & vbTab & "</Editors>"
                            End If
                        End If
                        '
                        Dim addonGuid As String = CS.GetText("ccGuid")
                        If (String.IsNullOrWhiteSpace(addonGuid)) Then
                            addonGuid = cp.Utils.CreateGuid()
                            Call CS.SetField("ccGuid", addonGuid)
                        End If
                        Dim NavType As String = CS.GetText("NavTypeID")
                        If (NavType = "") Then
                            NavType = "Add-on"
                        End If
                        result = "" _
                        & vbCrLf & vbTab & "<Addon name=""" & System.Net.WebUtility.HtmlEncode(addonName) & """ guid=""" & addonGuid & """ type=""" & NavType & """>" _
                        & tabIndent(cp, result) _
                        & vbCrLf & vbTab & "</Addon>"
                    End If
                    Call CS.Close()

                End Using
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
        Private Function getNodeText(cp As CPBaseClass, NodeName As String, NodeContent As String, Optional deprecated As Boolean = False) As String
            getNodeText = ""
            Try
                Dim prefix As String = ""
                If (deprecated) Then
                    prefix = "<!-- deprecated -->"
                End If
                getNodeText = ""
                If NodeContent = "" Then
                    getNodeText = getNodeText & vbCrLf & vbTab & prefix & "<" & NodeName & "></" & NodeName & ">"
                Else
                    getNodeText = getNodeText & vbCrLf & vbTab & prefix & "<" & NodeName & ">" & EncodeCData(cp, NodeContent) & "</" & NodeName & ">"
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
        Private Function getNodeBoolean(cp As CPBaseClass, NodeName As String, NodeContent As Boolean) As String
            getNodeBoolean = ""
            Try
                getNodeBoolean = vbCrLf & vbTab & "<" & NodeName & ">" & getYesNo(cp, NodeContent) & "</" & NodeName & ">"
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
                Call Doc.LoadXml(cp.PrivateFiles.Read("addons\Collections.xml"))
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
            Catch ex As Exception
                errorReport(cp, ex, "GetLocalCollectionArgs")
            End Try
        End Sub
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
        Public Function getYesNo(cp As CPBaseClass, Key As Boolean) As String
            Return If(Key, "Yes", "No")
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
        ''' <param name="zipTempPathFilename"></param>
        ''' <param name="addTempPathFilename"></param>
        ''' <remarks></remarks>
        Public Sub zipTempCdnFile(cp As CPBaseClass, zipTempPathFilename As String, ByVal addTempPathFilename As List(Of String))
            Try
                Dim z As Zip.ZipFile
                If cp.TempFiles.FileExists(zipTempPathFilename) Then
                    '
                    ' update existing zip with list of files
                    z = New Zip.ZipFile(cp.TempFiles.PhysicalFilePath & zipTempPathFilename)
                Else
                    '
                    ' create new zip
                    z = Zip.ZipFile.Create(cp.TempFiles.PhysicalFilePath & zipTempPathFilename)
                End If
                z.BeginUpdate()
                For Each pathFilename In addTempPathFilename
                    z.Add(cp.TempFiles.PhysicalFilePath & pathFilename, System.IO.Path.GetFileName(pathFilename))
                Next
                z.CommitUpdate()
                z.Close()
            Catch ex As Exception
                errorReport(cp, ex, "zipFile")
            End Try
        End Sub
        '
        '=======================================================================================
        '
        Private Function getPath(cp As CPBaseClass, ByVal pathFilename As String) As String
            Dim Position As Integer = InStrRev(pathFilename, "\")
            If Position <> 0 Then
                Return Mid(pathFilename, 1, Position)
            End If
            Return String.Empty
        End Function
        '
        '=======================================================================================
        '
        Public Function getFilename(cp As CPBaseClass, ByVal PathFilename As String) As String
            Dim pos As Integer = InStrRev(PathFilename, "/")
            If pos <> 0 Then
                Return Mid(PathFilename, pos + 1)
            End If
            Return PathFilename
        End Function
        '
        '=======================================================================================
        '
        '   Indent every line by 1 tab
        '
        Public Function tabIndent(cp As CPBaseClass, Source As String) As String
            Dim posStart As Integer = InStr(1, Source, "<![CDATA[", CompareMethod.Text)
            If posStart = 0 Then
                '
                ' no cdata
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
                    Dim posEnd As Integer = InStr(posStart, Source, "</textarea>", CompareMethod.Text)
                    Dim pre As String = Mid(Source, 1, posStart - 1)
                    Dim post As String = ""
                    Dim target As String
                    If posEnd = 0 Then
                        target = Mid(Source, posStart)
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
                Dim posEnd As Integer = InStr(posStart, Source, "]]>", CompareMethod.Text)
                Dim pre As String = Mid(Source, 1, posStart - 1)
                Dim post As String = ""
                Dim target As String
                If posEnd = 0 Then
                    target = Mid(Source, posStart)
                Else
                    target = Mid(Source, posStart, posEnd - posStart + Len("]]>"))
                    post = Mid(Source, posEnd + 3)
                End If
                tabIndent = tabIndent(cp, pre) & target & tabIndent(cp, post)
            End If
        End Function
    End Class
End Namespace
