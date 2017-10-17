
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.addonManager
    Public Class xmlController
        '
        '========================================================================
        ' This page and its contents are copyright by Kidwell McGowan Associates.
        '========================================================================
        '
        ' ----- global scope variables
        '
        Private iAbort As Boolean
        Private iBusy As Integer
        Private iTaskCount As Integer
        Const ApplicationNameLocal = "unknown"
        Private cp As CPBaseClass
        '
        ' ----- This should match the Lookup List in the NavIconType field in the Navigator Entry content definition
        '
        Public Const navTypeIDList = "Add-on,Report,Setting,Tool"
        Public Const NavTypeIDAddon = 1
        Public Const NavTypeIDReport = 2
        Public Const NavTypeIDSetting = 3
        Public Const NavTypeIDTool = 4
        '
        Public Const NavIconTypeList = "Custom,Advanced,Content,Folder,Email,User,Report,Setting,Tool,Record,Addon,help"
        Public Const NavIconTypeCustom = 1
        Public Const NavIconTypeAdvanced = 2
        Public Const NavIconTypeContent = 3
        Public Const NavIconTypeFolder = 4
        Public Const NavIconTypeEmail = 5
        Public Const NavIconTypeUser = 6
        Public Const NavIconTypeReport = 7
        Public Const NavIconTypeSetting = 8
        Public Const NavIconTypeTool = 9
        Public Const NavIconTypeRecord = 10
        Public Const NavIconTypeAddon = 11
        Public Const NavIconTypeHelp = 12
        '
        Public Const FieldDescriptorInteger = "Integer"
        Public Const FieldDescriptorText = "Text"
        Public Const FieldDescriptorLongText = "LongText"
        Public Const FieldDescriptorBoolean = "Boolean"
        Public Const FieldDescriptorDate = "Date"
        Public Const FieldDescriptorFile = "File"
        Public Const FieldDescriptorLookup = "Lookup"
        Public Const FieldDescriptorRedirect = "Redirect"
        Public Const FieldDescriptorCurrency = "Currency"
        Public Const FieldDescriptorImage = "Image"
        Public Const FieldDescriptorFloat = "Float"
        Public Const FieldDescriptorManyToMany = "ManyToMany"
        Public Const FieldDescriptorTextFile = "TextFile"
        Public Const FieldDescriptorCSSFile = "CSSFile"
        Public Const FieldDescriptorXMLFile = "XMLFile"
        Public Const FieldDescriptorJavascriptFile = "JavascriptFile"
        Public Const FieldDescriptorLink = "Link"
        Public Const FieldDescriptorResourceLink = "ResourceLink"
        Public Const FieldDescriptorMemberSelect = "MemberSelect"
        Public Const FieldDescriptorHTML = "HTML"
        Public Const FieldDescriptorHTMLFile = "HTMLFile"
        '
        Public Const FieldDescriptorLcaseInteger = "integer"
        Public Const FieldDescriptorLcaseText = "text"
        Public Const FieldDescriptorLcaseLongText = "longtext"
        Public Const FieldDescriptorLcaseBoolean = "boolean"
        Public Const FieldDescriptorLcaseDate = "date"
        Public Const FieldDescriptorLcaseFile = "file"
        Public Const FieldDescriptorLcaseLookup = "lookup"
        Public Const FieldDescriptorLcaseRedirect = "redirect"
        Public Const FieldDescriptorLcaseCurrency = "currency"
        Public Const FieldDescriptorLcaseImage = "image"
        Public Const FieldDescriptorLcaseFloat = "float"
        Public Const FieldDescriptorLcaseManyToMany = "manytomany"
        Public Const FieldDescriptorLcaseTextFile = "textfile"
        Public Const FieldDescriptorLcaseCSSFile = "cssfile"
        Public Const FieldDescriptorLcaseXMLFile = "xmlfile"
        Public Const FieldDescriptorLcaseJavascriptFile = "javascriptfile"
        Public Const FieldDescriptorLcaseLink = "link"
        Public Const FieldDescriptorLcaseResourceLink = "resourcelink"
        Public Const FieldDescriptorLcaseMemberSelect = "memberselect"
        Public Const FieldDescriptorLcaseHTML = "html"
        Public Const FieldDescriptorLcaseHTMLFile = "htmlfile"
        '
        '====================================================================================================
        ''' <summary>
        ''' constructor
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <remarks></remarks>
        Public Sub New(cp As CPBaseClass)
            Me.cp = cp
        End Sub
        '
        '====================================================================================================
        Private Class tableClass
            Public tableName As String
            Public dataSourceName As String
        End Class
        '
        '====================================================================================================
        Public Function GetXMLContentDefinition3(Optional ByVal ContentName As String = "", Optional ByVal IncludeBaseFields As Boolean = False) As String
            GetXMLContentDefinition3 = ""
            Try
                '
                Const ContentSelectList = "" _
                    & " id,name,active,adminonly,allowadd" _
                    & ",allowcalendarevents,allowcontentchildtool,allowcontenttracking,allowdelete,allowmetacontent" _
                    & ",allowtopicrules,AllowWorkflowAuthoring,AuthoringTableID" _
                    & ",ContentTableID,DefaultSortMethodID,DeveloperOnly,DropDownFieldList" _
                    & ",EditorGroupID,ParentID,ccGuid,IsBaseContent" _
                    & ",IconLink,IconHeight,IconWidth,IconSprites"

                Const FieldSelectList = "" _
                    & "f.ID,f.Name,f.contentid,f.Active,f.AdminOnly,f.Authorable,f.Caption,f.DeveloperOnly,f.EditSortPriority,f.Type,f.HTMLContent" _
                    & ",f.IndexColumn,f.IndexSortDirection,f.IndexSortPriority,f.RedirectID,f.RedirectPath,f.Required" _
                    & ",f.TextBuffered,f.UniqueName,f.DefaultValue,f.RSSTitleField,f.RSSDescriptionField,f.MemberSelectGroupID" _
                    & ",f.EditTab,f.Scramble,f.LookupList,f.NotEditable,f.Password,f.readonly,f.ManyToManyRulePrimaryField" _
                    & ",f.ManyToManyRuleSecondaryField,'' as HelpMessageDeprecated,f.ModifiedBy,f.IsBaseField,f.LookupContentID" _
                    & ",f.RedirectContentID,f.ManyToManyContentID,f.ManyToManyRuleContentID" _
                    & ",h.helpdefault,h.helpcustom,f.IndexWidth"

                '
                Dim IsBaseContent As Boolean
                Dim FieldCnt As Integer
                Dim FieldName As String
                Dim FieldContentID As Integer
                Dim LastFieldID As Integer
                Dim RecordID As Integer
                Dim RecordName As String
                Dim AuthoringTableID As Integer
                Dim HelpDefault As String
                Dim HelpCnt As Integer
                Dim fieldId As Integer
                Dim fieldType As String
                Dim ContentTableID As Integer
                Dim TableName As String
                Dim DataSourceName As String
                Dim DefaultSortMethodID As Integer
                Dim DefaultSortMethod As String
                Dim EditorGroupID As Integer
                Dim EditorGroupName As String
                Dim ParentID As Integer
                Dim ParentName As String
                Dim ContentID As Integer
                Dim sb As New System.Text.StringBuilder
                Dim iContentName As String
                Dim SQL As String
                Dim FoundMenuTable As Boolean
                Dim appName As String
                Dim cs As CPCSBaseClass = cp.CSNew()
                '
                appName = cp.Site.Name
                iContentName = ContentName
                If iContentName <> "" Then
                    SQL = "select id from cccontent where name=" & cp.Db.EncodeSQLText(iContentName)
                    If cs.OpenSQL(SQL) Then
                        ContentID = cs.GetInteger("id")
                    End If
                    Call cs.Close()
                End If
                If iContentName <> "" And (ContentID = 0) Then
                    '
                    ' export requested for content name that does not exist - return blank
                    '
                Else
                    '
                    ' Build table lookup
                    '
                    Dim tables As New Dictionary(Of Integer, tableClass)
                    SQL = "select T.ID,T.Name as TableName,D.Name as DataSourceName from ccTables T Left Join ccDataSources D on D.ID=T.DataSourceID"
                    If cs.OpenSQL(SQL) Then
                        Do
                            Dim table As New tableClass
                            table.tableName = cs.GetText("TableName")
                            table.dataSourceName = cs.GetText("DataSourceName")
                            tables.Add(cs.GetInteger("id"), table)
                            Call cs.GoNext()
                        Loop While cs.OK()
                    End If
                    Call cs.Close()
                    '
                    '
                    ' Build SortMethod lookup
                    '
                    Dim sorts As New Dictionary(Of Integer, String)
                    SQL = "select ID,Name from ccSortMethods"
                    If cs.OpenSQL(SQL) Then
                        Do
                            sorts.Add(cs.GetInteger("id"), cs.GetText("name"))
                            Call cs.GoNext()
                        Loop While cs.OK()
                    End If
                    Call cs.Close()
                    '
                    ' Build groups lookup
                    '
                    Dim groups As New Dictionary(Of Integer, String)
                    SQL = "select ID,Name from ccGroups"
                    If cs.OpenSQL(SQL) Then
                        Do
                            groups.Add(cs.GetInteger("id"), cs.GetText("name"))
                            Call cs.GoNext()
                        Loop While cs.OK()
                    End If
                    Call cs.Close()
                    '
                    ' Build Content lookup
                    '
                    SQL = "select id,name from ccContent"
                    Dim contents As New Dictionary(Of Integer, String)
                    If cs.OpenSQL(SQL) Then
                        Do
                            contents.Add(cs.GetInteger("id"), cs.GetText("name"))
                            Call cs.GoNext()
                        Loop While cs.OK()
                    End If
                    Call cs.Close()
                    ''
                    '' select all the fields
                    ''
                    'If ContentID <> 0 Then
                    '    SQL = "select " & FieldSelectList & "" _
                    '        & " from ccfields f left join ccfieldhelp h on h.fieldid=f.id" _
                    '        & " where (f.Type<>0)and(f.contentid=" & ContentID & ")" _
                    '        & ""
                    'Else
                    '    SQL = "select " & FieldSelectList & "" _
                    '        & " from ccfields f left join ccfieldhelp h on h.fieldid=f.id" _
                    '        & " where (f.Type<>0)" _
                    '        & ""
                    'End If
                    'If Not IncludeBaseFields Then
                    '    SQL = SQL & " and ((f.IsBaseField is null)or(f.IsBaseField=0))"
                    'End If
                    'SQL = SQL & " order by f.contentid,f.id,h.id desc"

                    'RS = cpCore.app.executeSql(SQL)
                    'CFields = convertDataTabletoArray(RS)
                    'CFieldCnt = UBound(CFields, 2) + 1
                    '
                    ' select the content
                    '
                    If ContentID <> 0 Then
                        SQL = "select " & ContentSelectList & " from ccContent where (id=" & ContentID & ")and(contenttableid is not null)and(contentcontrolid is not null) order by id"
                    Else
                        SQL = "select " & ContentSelectList & " from ccContent where (name<>'')and(name is not null)and(contenttableid is not null)and(contentcontrolid is not null) order by id"
                    End If
                    Dim csContent As CPCSBaseClass = cp.CSNew()
                    If csContent.OpenSQL(SQL) Then
                        '
                        ' ----- <cdef>
                        '
                        IsBaseContent = (csContent.GetBoolean("isBaseContent"))
                        iContentName = GetRSXMLAttribute(csContent, "Name")
                        If InStr(1, iContentName, "data sources", vbTextCompare) = 1 Then
                            iContentName = iContentName
                        End If
                        ContentID = (csContent.GetInteger("ID"))
                        sb.Append(vbCrLf & vbTab & "<CDef")
                        sb.Append(" Name=""" & iContentName & """")
                        If (Not IsBaseContent) Or IncludeBaseFields Then
                            sb.Append(" Active=""" & GetRSXMLAttribute(csContent, "Active") & """")
                            sb.Append(" AdminOnly=""" & GetRSXMLAttribute(csContent, "AdminOnly") & """")
                            'sb.Append( " AliasID=""" & GetRSXMLAttribute( appname,RS, "AliasID") & """")
                            'sb.Append( " AliasName=""" & GetRSXMLAttribute( appname,RS, "AliasName") & """")
                            sb.Append(" AllowAdd=""" & GetRSXMLAttribute(csContent, "AllowAdd") & """")
                            sb.Append(" AllowCalendarEvents=""" & GetRSXMLAttribute(csContent, "AllowCalendarEvents") & """")
                            sb.Append(" AllowContentChildTool=""" & GetRSXMLAttribute(csContent, "AllowContentChildTool") & """")
                            sb.Append(" AllowContentTracking=""" & GetRSXMLAttribute(csContent, "AllowContentTracking") & """")
                            sb.Append(" AllowDelete=""" & GetRSXMLAttribute(csContent, "AllowDelete") & """")
                            sb.Append(" AllowMetaContent=""" & GetRSXMLAttribute(csContent, "AllowMetaContent") & """")
                            sb.Append(" AllowTopicRules=""" & GetRSXMLAttribute(csContent, "AllowTopicRules") & """")
                            sb.Append(" AllowWorkflowAuthoring=""" & GetRSXMLAttribute(csContent, "AllowWorkflowAuthoring") & """")
                            '
                            AuthoringTableID = (csContent.GetInteger("AuthoringTableID"))
                            TableName = ""
                            DataSourceName = ""
                            If (tables.ContainsKey(AuthoringTableID)) Then
                                TableName = tables(AuthoringTableID).tableName
                                DataSourceName = tables(AuthoringTableID).dataSourceName
                            End If
                            If DataSourceName = "" Then
                                DataSourceName = "Default"
                            End If
                            If UCase(TableName) = "CCMENUENTRIES" Then
                                FoundMenuTable = True
                            End If
                            sb.Append(" AuthoringDataSourceName=""" & EncodeXMLattribute(DataSourceName) & """")
                            sb.Append(" AuthoringTableName=""" & EncodeXMLattribute(TableName) & """")
                            '
                            ContentTableID = (csContent.GetInteger("ContentTableID"))
                            If ContentTableID <> AuthoringTableID Then
                                If ContentTableID <> 0 Then
                                    TableName = ""
                                    DataSourceName = ""
                                    If (tables.ContainsKey(ContentTableID)) Then
                                        TableName = tables(ContentTableID).tableName
                                        DataSourceName = tables(ContentTableID).dataSourceName
                                        If DataSourceName = "" Then
                                            DataSourceName = "Default"
                                        End If
                                    End If
                                End If
                            End If
                            sb.Append(" ContentDataSourceName=""" & EncodeXMLattribute(DataSourceName) & """")
                            sb.Append(" ContentTableName=""" & EncodeXMLattribute(TableName) & """")
                            '
                            DefaultSortMethod = ""
                            DefaultSortMethodID = (csContent.GetInteger("DefaultSortMethodID"))
                            If (sorts.ContainsKey(DefaultSortMethodID)) Then
                                DefaultSortMethod = sorts(DefaultSortMethodID)
                            End If
                            sb.Append(" DefaultSortMethod=""" & EncodeXMLattribute(DefaultSortMethod) & """")
                            '
                            sb.Append(" DeveloperOnly=""" & GetRSXMLAttribute(csContent, "DeveloperOnly") & """")
                            sb.Append(" DropDownFieldList=""" & GetRSXMLAttribute(csContent, "DropDownFieldList") & """")
                            '
                            EditorGroupName = ""
                            EditorGroupID = (csContent.GetInteger("EditorGroupID"))
                            If (groups.ContainsKey(EditorGroupID)) Then
                                EditorGroupName = groups(EditorGroupID)
                            End If
                            sb.Append(" EditorGroupName=""" & EncodeXMLattribute(EditorGroupName) & """")
                            '
                            ParentName = ""
                            ParentID = (csContent.GetInteger("ParentID"))
                            If (contents.ContainsKey(ParentID)) Then
                                ParentName = contents(ParentID)
                            End If
                            sb.Append(" Parent=""" & EncodeXMLattribute(ParentName) & """")
                            '
                            sb.Append(" IconLink=""" & GetRSXMLAttribute(csContent, "IconLink") & """")
                            sb.Append(" IconHeight=""" & GetRSXMLAttribute(csContent, "IconHeight") & """")
                            sb.Append(" IconWidth=""" & GetRSXMLAttribute(csContent, "IconWidth") & """")
                            sb.Append(" IconSprites=""" & GetRSXMLAttribute(csContent, "IconSprites") & """")
                            '
                            '
                            If True Then
                                '
                                ' Add IsBaseContent
                                '
                                sb.Append(" isbasecontent=""" & GetRSXMLAttribute(csContent, "IsBaseContent") & """")
                            End If
                        End If
                        '
                        If True Then
                            '
                            ' Add guid
                            '
                            sb.Append(" guid=""" & GetRSXMLAttribute(csContent, "ccGuid") & """")
                        End If
                        sb.Append(" >")
                        '
                        ' create output
                        '
                        If ContentID <> 0 Then
                            SQL = "select " & FieldSelectList & "" _
                                & " from ccfields f left join ccfieldhelp h on h.fieldid=f.id" _
                                & " where (f.Type<>0)and(f.contentid=" & ContentID & ")" _
                                & ""
                        Else
                            SQL = "select " & FieldSelectList & "" _
                                & " from ccfields f left join ccfieldhelp h on h.fieldid=f.id" _
                                & " where (f.Type<>0)" _
                                & ""
                        End If
                        If Not IncludeBaseFields Then
                            SQL = SQL & " and ((f.IsBaseField is null)or(f.IsBaseField=0))"
                        End If
                        SQL = SQL & " order by f.contentid,f.id,h.id desc"
                        Dim CFields As CPCSBaseClass = cp.CSNew()
                        If (CFields.OpenSQL(SQL)) Then
                            fieldId = 0
                            Do

                                LastFieldID = fieldId
                                fieldId = CFields.GetInteger("ID")
                                FieldName = CFields.GetText("Name")
                                FieldContentID = CFields.GetInteger("contentid")
                                If FieldContentID > ContentID Then
                                    Exit Do
                                ElseIf (FieldContentID = ContentID) And (fieldId <> LastFieldID) Then
                                    If IncludeBaseFields Or (InStr(1, ",id,ContentCategoryID,dateadded,createdby,modifiedby,EditBlank,EditArchive,EditSourceID,ContentControlID,CreateKey,ModifiedDate,ccguid,", "," & FieldName & ",", vbTextCompare) = 0) Then
                                        sb.Append(vbCrLf & vbTab & vbTab & "<Field")
                                        fieldType = csv_GetFieldDescriptorByType(CFields.GetInteger("Type"))
                                        sb.Append(" Name=""" & FieldName & """")
                                        sb.Append(" active=""" & CFields.GetBoolean("Active") & """")
                                        sb.Append(" AdminOnly=""" & CFields.GetBoolean("AdminOnly") & """")
                                        sb.Append(" Authorable=""" & CFields.GetBoolean("Authorable") & """")
                                        sb.Append(" Caption=""" & CFields.GetText("Caption") & """")
                                        sb.Append(" DeveloperOnly=""" & CFields.GetBoolean("DeveloperOnly") & """")
                                        sb.Append(" EditSortPriority=""" & CFields.GetText("EditSortPriority") & """")
                                        sb.Append(" FieldType=""" & fieldType & """")
                                        sb.Append(" HTMLContent=""" & CFields.GetBoolean("HTMLContent") & """")
                                        sb.Append(" IndexColumn=""" & CFields.GetText("IndexColumn") & """")
                                        sb.Append(" IndexSortDirection=""" & CFields.GetText("IndexSortDirection") & """")
                                        sb.Append(" IndexSortOrder=""" & CFields.GetText("IndexSortPriority") & """")
                                        sb.Append(" IndexWidth=""" & CFields.GetText("IndexWidth") & """")
                                        sb.Append(" RedirectID=""" & CFields.GetText("RedirectID") & """")
                                        sb.Append(" RedirectPath=""" & CFields.GetText("RedirectPath") & """")
                                        sb.Append(" Required=""" & CFields.GetBoolean("Required") & """")
                                        sb.Append(" TextBuffered=""" & CFields.GetBoolean("TextBuffered") & """")
                                        sb.Append(" UniqueName=""" & CFields.GetBoolean("UniqueName") & """")
                                        sb.Append(" DefaultValue=""" & CFields.GetText("DefaultValue") & """")
                                        sb.Append(" RSSTitle=""" & CFields.GetBoolean("RSSTitleField") & """")
                                        sb.Append(" RSSDescription=""" & CFields.GetBoolean("RSSDescriptionField") & """")
                                        sb.Append(" MemberSelectGroupID=""" & CFields.GetText("MemberSelectGroupID") & """")
                                        sb.Append(" EditTab=""" & CFields.GetText("EditTab") & """")
                                        sb.Append(" Scramble=""" & CFields.GetBoolean("Scramble") & """")
                                        sb.Append(" LookupList=""" & CFields.GetText("LookupList") & """")
                                        sb.Append(" NotEditable=""" & CFields.GetBoolean("NotEditable") & """")
                                        sb.Append(" Password=""" & CFields.GetBoolean("Password") & """")
                                        sb.Append(" ReadOnly=""" & CFields.GetBoolean("ReadOnly") & """")
                                        sb.Append(" ManyToManyRulePrimaryField=""" & CFields.GetText("ManyToManyRulePrimaryField") & """")
                                        sb.Append(" ManyToManyRuleSecondaryField=""" & CFields.GetText("ManyToManyRuleSecondaryField") & """")
                                        sb.Append(" IsModified=""" & (CFields.GetInteger("ModifiedBy") <> 0) & """")
                                        If True Then
                                            sb.Append(" IsBaseField=""" & CFields.GetBoolean("IsBaseField") & """")
                                        End If
                                        '
                                        RecordName = ""
                                        RecordID = CFields.GetInteger("LookupContentID")
                                        If (contents.ContainsKey(RecordID)) Then
                                            RecordName = contents(RecordID)
                                        End If
                                        sb.Append(" LookupContent=""" & cp.Utils.EncodeHTML(RecordName) & """")
                                        '
                                        RecordName = ""
                                        RecordID = CFields.GetInteger("RedirectContentID")
                                        If (contents.ContainsKey(RecordID)) Then
                                            RecordName = contents(RecordID)
                                        End If
                                        sb.Append(" RedirectContent=""" & cp.Utils.EncodeHTML(RecordName) & """")
                                        '
                                        RecordName = ""
                                        RecordID = CFields.GetInteger("ManyToManyContentID")
                                        If (contents.ContainsKey(RecordID)) Then
                                            RecordName = contents(RecordID)
                                        End If
                                        sb.Append(" ManyToManyContent=""" & cp.Utils.EncodeHTML(RecordName) & """")
                                        '
                                        RecordName = ""
                                        RecordID = CFields.GetInteger("ManyToManyRuleContentID")
                                        If (contents.ContainsKey(RecordID)) Then
                                            RecordName = contents(RecordID)
                                        End If
                                        sb.Append(" ManyToManyRuleContent=""" & cp.Utils.EncodeHTML(RecordName) & """")
                                        '
                                        sb.Append(" >")
                                        '
                                        HelpCnt = 0
                                        HelpDefault = CFields.GetText("helpcustom")
                                        If HelpDefault = "" Then
                                            HelpDefault = CFields.GetText("helpdefault")
                                        End If
                                        If HelpDefault <> "" Then
                                            sb.Append(vbCrLf & vbTab & vbTab & vbTab & "<HelpDefault>" & EncodeCData(HelpDefault) & "</HelpDefault>")
                                            HelpCnt = HelpCnt + 1
                                        End If
                                        '                            HelpCustom = cfields.getText("helpcustom")
                                        '                            If HelpCustom <> "" Then
                                        '                                sb.Append( vbCrLf & vbTab & vbTab & vbTab & "<HelpCustom>" & HelpCustom & "</HelpCustom>")
                                        '                                HelpCnt = HelpCnt + 1
                                        '                            End If
                                        If HelpCnt > 0 Then
                                            sb.Append(vbCrLf & vbTab & vbTab)
                                        End If
                                        sb.Append("</Field>")
                                    End If
                                    FieldCnt = FieldCnt + 1
                                End If
                                Call CFields.GoNext()
                            Loop While CFields.OK()
                        End If
                        Call CFields.Close()
                        '
                        If FieldCnt > 0 Then
                            sb.Append(vbCrLf & vbTab)
                        End If
                        sb.Append("</CDef>")
                    End If
                    Call csContent.Close()
                    If ContentName = "" Then
                        '
                        ' Add other areas of the CDef file
                        '
                        sb.Append(GetXMLContentDefinition_SQLIndexes())
                        If FoundMenuTable Then
                            sb.Append(GetXMLContentDefinition_AdminMenus())
                        End If
                        '
                        ' These are not needed anymore - later add "ImportCollection" entries for all collections installed
                        '
                        '        If FoundAFTable Then
                        '            sb.Append( GetXMLContentDefinition_AggregateFunctions()
                        '        End If
                    End If
                    Const ApplicationCollectionGuid = "{C58A76E2-248B-4DE8-BF9C-849A960F79C6}"
                    Const CollectionFileRootNode = "collection"
                    GetXMLContentDefinition3 = "<" & CollectionFileRootNode & " name=""Application"" guid=""" & ApplicationCollectionGuid & """>" & sb.ToString & vbCrLf & "</" & CollectionFileRootNode & ">"
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
        End Function
        '
        '========================================================================
        ' ----- Save the admin menus to CDef AdminMenu tags
        '========================================================================
        '
        Private Function GetXMLContentDefinition_SQLIndexes() As String
            GetXMLContentDefinition_SQLIndexes = ""
            Try
                Dim DataSourceName As String
                Dim TableName As String
                '
                Dim IndexFields As String = ""
                Dim IndexList As String = ""
                Dim IndexName As String
                Dim ListRows() As String
                Dim ListRow As String = ""
                Dim ListRowSplit() As String
                Dim SQL As String
                Dim cs As CPCSBaseClass = cp.CSNew()
                Dim Ptr As Integer
                Dim sb As New System.Text.StringBuilder
                '
                SQL = "select D.name as DataSourceName,T.name as TableName" _
                    & " from cctables T left join ccDataSources d on D.ID=T.DataSourceID" _
                    & " where t.active<>0"
                If cs.OpenSQL(SQL) Then
                    Do
                        DataSourceName = cs.GetText("DataSourceName")
                        TableName = cs.GetText("TableName")
                        '
                        ' need a solution for this
                        '
                        'IndexList = cpCore.app.csv_GetSQLIndexList(DataSourceName, TableName)
                        '
                        If IndexList <> "" Then
                            ListRows = Split(IndexList, vbCrLf)
                            IndexName = ""
                            For Ptr = 0 To UBound(ListRows) + 1
                                If Ptr <= UBound(ListRows) Then
                                    '
                                    ' ListRowSplit has the indexname and field for this index
                                    '
                                    ListRowSplit = Split(ListRows(Ptr), ",")
                                Else
                                    '
                                    ' one past the last row, ListRowSplit gets a dummy entry to force the output of the last line
                                    '
                                    ListRowSplit = Split("-,-", ",")
                                End If
                                If UBound(ListRowSplit) > 0 Then
                                    If ListRowSplit(0) <> "" Then
                                        If IndexName = "" Then
                                            '
                                            ' first line of the first index description
                                            '
                                            IndexName = ListRowSplit(0)
                                            IndexFields = ListRowSplit(1)
                                        ElseIf IndexName = ListRowSplit(0) Then
                                            '
                                            ' next line of the index description
                                            '
                                            IndexFields = IndexFields & "," & ListRowSplit(1)
                                        Else
                                            '
                                            ' first line of a new index description
                                            ' save previous line
                                            '
                                            If IndexName <> "" And IndexFields <> "" Then
                                                Call sb.Append("<SQLIndex")
                                                Call sb.Append(" Indexname=""" & EncodeXMLattribute(IndexName) & """")
                                                Call sb.Append(" DataSourceName=""" & EncodeXMLattribute(DataSourceName) & """")
                                                Call sb.Append(" TableName=""" & EncodeXMLattribute(TableName) & """")
                                                Call sb.Append(" FieldNameList=""" & EncodeXMLattribute(IndexFields) & """")
                                                Call sb.Append("></SQLIndex>" & vbCrLf)
                                            End If
                                            '
                                            IndexName = ListRowSplit(0)
                                            IndexFields = ListRowSplit(1)
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        cs.GoNext()
                    Loop While cs.OK()
                End If
                cs.Close()
                GetXMLContentDefinition_SQLIndexes = sb.ToString
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
            '
        End Function
        '
        '========================================================================
        ' ----- Save the admin menus to CDef AdminMenu tags
        '========================================================================
        '
        Private Function GetXMLContentDefinition_AdminMenus() As String
            GetXMLContentDefinition_AdminMenus = ""
            Try
                Dim s As String = ""
                s = s & GetXMLContentDefinition_AdminMenus_MenuEntries()
                s = s & GetXMLContentDefinition_AdminMenus_NavigatorEntries()
                GetXMLContentDefinition_AdminMenus = s
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
        End Function
        '
        '========================================================================
        ' ----- Save the admin menus to CDef AdminMenu tags
        '========================================================================
        '
        Private Function GetXMLContentDefinition_AdminMenus_NavigatorEntries() As String
            GetXMLContentDefinition_AdminMenus_NavigatorEntries = ""
            Try
                Dim NavIconType As Integer
                Dim NavIconTitle As String
                Dim sb As New System.Text.StringBuilder
                Dim dt As CPCSBaseClass = cp.CSNew()
                Dim menuNameSpace As String
                Dim RecordName As String
                Dim ParentID As Integer
                Dim MenuContentID As Integer
                Dim SplitArray() As String
                Dim SplitIndex As Integer
                '
                ' ****************************** if cdef not loaded, this fails
                '
                MenuContentID = cp.Content.GetRecordID("Content", "Navigator Entries")
                If dt.OpenSQL("select * from ccMenuEntries where (contentcontrolid=" & MenuContentID & ")and(name<>'')") Then
                    NavIconType = 0
                    NavIconTitle = ""
                    Do
                        RecordName = (dt.GetText("Name"))
                        If RecordName = "Advanced" Then
                            RecordName = RecordName
                        End If
                        ParentID = (dt.GetInteger("ParentID"))
                        menuNameSpace = getMenuNameSpace(ParentID, "")
                        Call sb.Append("<NavigatorEntry Name=""" & EncodeXMLattribute(RecordName) & """")
                        Call sb.Append(" NameSpace=""" & menuNameSpace & """")
                        Call sb.Append(" LinkPage=""" & GetRSXMLAttribute(dt, "LinkPage") & """")
                        Call sb.Append(" ContentName=""" & GetRSXMLLookupAttribute(dt, "ContentID", "ccContent") & """")
                        Call sb.Append(" AdminOnly=""" & GetRSXMLAttribute(dt, "AdminOnly") & """")
                        Call sb.Append(" DeveloperOnly=""" & GetRSXMLAttribute(dt, "DeveloperOnly") & """")
                        Call sb.Append(" NewWindow=""" & GetRSXMLAttribute(dt, "NewWindow") & """")
                        Call sb.Append(" Active=""" & GetRSXMLAttribute(dt, "Active") & """")
                        Call sb.Append(" AddonName=""" & GetRSXMLLookupAttribute(dt, "AddonID", "ccAggregateFunctions") & """")
                        Call sb.Append(" SortOrder=""" & GetRSXMLAttribute(dt, "SortOrder") & """")
                        NavIconType = cp.Utils.EncodeInteger(GetRSXMLAttribute(dt, "NavIconType"))
                        NavIconTitle = GetRSXMLAttribute(dt, "NavIconTitle")
                        Call sb.Append(" NavIconTitle=""" & NavIconTitle & """")
                        SplitArray = Split(NavIconTypeList & ",help", ",")
                        SplitIndex = NavIconType - 1
                        If (SplitIndex >= 0) And (SplitIndex <= UBound(SplitArray)) Then
                            Call sb.Append(" NavIconType=""" & SplitArray(SplitIndex) & """")
                        Else
                            SplitIndex = SplitIndex
                        End If
                        '
                        If True Then
                            Call sb.Append(" guid=""" & GetRSXMLAttribute(dt, "ccGuid") & """")
                        ElseIf True Then
                            Call sb.Append(" guid=""" & GetRSXMLAttribute(dt, "NavGuid") & """")
                        End If
                        '
                        Call sb.Append("></NavigatorEntry>" & vbCrLf)
                        Call dt.GoNext()
                    Loop While dt.OK
                End If
                GetXMLContentDefinition_AdminMenus_NavigatorEntries = sb.ToString
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
            '
        End Function
        '
        '========================================================================
        ' ----- Save the admin menus to CDef AdminMenu tags
        '========================================================================
        '
        Private Function GetXMLContentDefinition_AdminMenus_MenuEntries() As String
            GetXMLContentDefinition_AdminMenus_MenuEntries = ""
            Try
                Dim sb As New StringBuilder
                Dim dr As CPCSBaseClass = cp.CSNew()
                Dim RecordName As String
                Dim MenuContentID As Integer
                '
                MenuContentID = cp.Content.GetRecordID("Content", "Menu Entries")
                If (dr.OpenSQL("select * from ccMenuEntries where (contentcontrolid=" & MenuContentID & ")and(name<>'')")) Then
                    Do
                        RecordName = (dr.GetText("Name"))
                        Call sb.Append("<MenuEntry Name=""" & EncodeXMLattribute(RecordName) & """")
                        Call sb.Append(" ParentName=""" & GetRSXMLLookupAttribute(dr, "ParentID", "ccMenuEntries") & """")
                        Call sb.Append(" LinkPage=""" & GetRSXMLAttribute(dr, "LinkPage") & """")
                        Call sb.Append(" ContentName=""" & GetRSXMLLookupAttribute(dr, "ContentID", "ccContent") & """")
                        Call sb.Append(" AdminOnly=""" & GetRSXMLAttribute(dr, "AdminOnly") & """")
                        Call sb.Append(" DeveloperOnly=""" & GetRSXMLAttribute(dr, "DeveloperOnly") & """")
                        Call sb.Append(" NewWindow=""" & GetRSXMLAttribute(dr, "NewWindow") & """")
                        Call sb.Append(" Active=""" & GetRSXMLAttribute(dr, "Active") & """")
                        If True Then
                            Call sb.Append(" AddonName=""" & GetRSXMLLookupAttribute(dr, "AddonID", "ccAggregateFunctions") & """")
                        End If
                        Call sb.Append("/>" & vbCrLf)
                        Call dr.GoNext()
                    Loop While dr.OK
                End If
                Call dr.Close()
                GetXMLContentDefinition_AdminMenus_MenuEntries = sb.ToString
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
        End Function
        '
        '========================================================================
        '
        '========================================================================
        '
        Private Function GetXMLContentDefinition_AggregateFunctions() As String
            GetXMLContentDefinition_AggregateFunctions = ""
            Try
                Dim rs As CPCSBaseClass = cp.CSNew()
                Dim sb As New System.Text.StringBuilder
                '
                If (rs.OpenSQL("select * from ccAggregateFunctions")) Then
                    Do
                        Call sb.Append("<Addon Name=""" & GetRSXMLAttribute(rs, "Name") & """")
                        Call sb.Append(" Link=""" & GetRSXMLAttribute(rs, "Link") & """")
                        Call sb.Append(" ObjectProgramID=""" & GetRSXMLAttribute(rs, "ObjectProgramID") & """")
                        Call sb.Append(" ArgumentList=""" & GetRSXMLAttribute(rs, "ArgumentList") & """")
                        Call sb.Append(" SortOrder=""" & GetRSXMLAttribute(rs, "SortOrder") & """")
                        Call sb.Append(" >")
                        Call sb.Append(GetRSXMLAttribute(rs, "Copy"))
                        Call sb.Append("</Addon>" & vbCrLf)
                        Call rs.GoNext()
                    Loop While rs.OK()
                End If
                GetXMLContentDefinition_AggregateFunctions = sb.ToString
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
        End Function
        '
        '
        '
        Private Function EncodeXMLattribute(ByVal Source As String) As String
            EncodeXMLattribute = cp.Utils.EncodeHTML(Source)
            EncodeXMLattribute = Replace(EncodeXMLattribute, vbCrLf, " ")
            EncodeXMLattribute = Replace(EncodeXMLattribute, vbCr, "")
            EncodeXMLattribute = Replace(EncodeXMLattribute, vbLf, "")
        End Function
        '
        '
        '
        Private Function GetTableRecordName(ByVal TableName As String, ByVal RecordID As Integer) As String
            GetTableRecordName = ""
            Try
                Dim dt As CPCSBaseClass = cp.CSNew()
                '
                If RecordID <> 0 And TableName <> "" Then
                    If dt.OpenSQL("select Name from " & TableName & " where ID=" & RecordID) Then
                        GetTableRecordName = dt.GetText("name")
                    End If
                    dt.Close()
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
        End Function
        '
        '
        '
        Private Function GetRSXMLAttribute(ByVal dr As CPCSBaseClass, ByVal FieldName As String) As String
            GetRSXMLAttribute = ""
            Try
                GetRSXMLAttribute = EncodeXMLattribute((dr.GetText(FieldName)))
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
        End Function
        '
        '
        '
        Private Function GetRSXMLLookupAttribute(ByVal dr As CPCSBaseClass, ByVal FieldName As String, ByVal TableName As String) As String
            GetRSXMLLookupAttribute = ""
            Try
                GetRSXMLLookupAttribute = EncodeXMLattribute(GetTableRecordName(TableName, (dr.GetInteger(FieldName))))
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
        End Function
        '
        '
        '
        Private Function getMenuNameSpace(ByVal RecordID As Integer, ByVal UsedIDString As String) As String
            getMenuNameSpace = ""
            Try
                Dim rs As CPCSBaseClass = cp.CSNew()
                Dim ParentID As Integer
                Dim RecordName As String = ""
                Dim ParentSpace As String = ""
                '
                If RecordID <> 0 Then
                    If InStr(1, "," & UsedIDString & ",", "," & RecordID & ",", vbTextCompare) <> 0 Then
                        cp.Site.ErrorReport("Circular reference found in UsedIDString [" & UsedIDString & "] getting ccMenuEntries namespace for recordid [" & RecordID & "]")
                        getMenuNameSpace = ""
                    Else
                        UsedIDString = UsedIDString & "," & RecordID
                        ParentID = 0
                        If RecordID <> 0 Then
                            If (rs.OpenSQL("select Name,ParentID from ccMenuEntries where ID=" & RecordID)) Then
                                ParentID = rs.GetInteger("parentid")
                                RecordName = rs.GetText("name")
                            End If
                            rs.Close()
                        End If
                        If RecordName <> "" Then
                            If ParentID = RecordID Then
                                '
                                ' circular reference
                                '
                                cp.Site.ErrorReport("Circular reference found (ParentID=RecordID) getting ccMenuEntries namespace for recordid [" & RecordID & "]")
                                getMenuNameSpace = ""
                            Else
                                If ParentID <> 0 Then
                                    '
                                    ' get next parent
                                    '
                                    ParentSpace = getMenuNameSpace(ParentID, UsedIDString)
                                End If
                                If ParentSpace <> "" Then
                                    getMenuNameSpace = ParentSpace & "." & RecordName
                                Else
                                    getMenuNameSpace = RecordName
                                End If
                            End If
                        Else
                            getMenuNameSpace = ""
                        End If
                    End If
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
        End Function
        '
        '========================================================================
        ' ----- Get FieldDescritor from FieldType
        '========================================================================
        '
        Public Function csv_GetFieldDescriptorByType(ByVal fieldType As Integer) As String
            Dim result As String = ""
            result = ""
            Try
                Select Case fieldType
                    Case exportClass.FieldTypeBoolean
                        result = FieldDescriptorBoolean
                    Case exportClass.FieldTypeCurrency
                        result = FieldDescriptorCurrency
                    Case exportClass.FieldTypeDate
                        result = FieldDescriptorDate
                    Case exportClass.FieldTypeFile
                        result = FieldDescriptorFile
                    Case exportClass.FieldTypeFloat
                        result = FieldDescriptorFloat
                    Case exportClass.FieldTypeImage
                        result = FieldDescriptorImage
                    Case exportClass.FieldTypeLink
                        result = FieldDescriptorLink
                    Case exportClass.FieldTypeResourceLink
                        result = FieldDescriptorResourceLink
                    Case exportClass.FieldTypeInteger
                        result = FieldDescriptorInteger
                    Case exportClass.FieldTypeLongText
                        result = FieldDescriptorLongText
                    Case exportClass.FieldTypeLookup
                        result = FieldDescriptorLookup
                    Case exportClass.FieldTypeMemberSelect
                        result = FieldDescriptorMemberSelect
                    Case exportClass.FieldTypeRedirect
                        result = FieldDescriptorRedirect
                    Case exportClass.FieldTypeManyToMany
                        result = FieldDescriptorManyToMany
                    Case exportClass.FieldTypeTextFile
                        result = FieldDescriptorTextFile
                    Case exportClass.FieldTypeCSSFile
                        result = FieldDescriptorCSSFile
                    Case exportClass.FieldTypeXMLFile
                        result = FieldDescriptorXMLFile
                    Case exportClass.FieldTypeJavascriptFile
                        result = FieldDescriptorJavascriptFile
                    Case exportClass.FieldTypeText
                        result = FieldDescriptorText
                    Case exportClass.FieldTypeHTML
                        result = FieldDescriptorHTML
                    Case exportClass.FieldTypeHTMLFile
                        result = FieldDescriptorHTMLFile
                    Case Else
                        If fieldType = exportClass.FieldTypeAutoIncrement Then
                            result = "AutoIncrement"
                        ElseIf fieldType = exportClass.FieldTypeMemberSelect Then
                            result = "MemberSelect"
                        Else
                            '
                            ' If field type is ignored, call it a text field
                            '
                            result = FieldDescriptorText
                        End If
                End Select
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetXMLContentDefinition3")
            End Try
            Return result
        End Function
        '
        '====================================================================================================
        Private Function EncodeCData(Source As String) As String
            EncodeCData = ""
            Try
                EncodeCData = Source
                If EncodeCData <> "" Then
                    EncodeCData = "<![CDATA[" & Replace(EncodeCData, "]]>", "]]]]><![CDATA[>") & "]]>"
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "EncodeCData")
            End Try
        End Function
    End Class
End Namespace