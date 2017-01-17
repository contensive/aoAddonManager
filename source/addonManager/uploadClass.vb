
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports Contensive.addonManager

Namespace Contensive.addonManager
    '
    ' Sample Vb addon
    '
    Public Class uploadClass
        Inherits AddonBaseClass
        '
        ' injected objects -- do not dispose
        '
        Private cp As CPBaseClass
        ' 
        ' class scope
        '
        Private Const RequestNameButton As String = "button"
        Private Const CollectionListRootNode As String = "collectionlist"
        '
        Private Const ButtonCancel As String = " Cancel "
        Private Const ButtonOK As String = " Cancel "
        '
        Private Structure NavigatorType
            Public Name As String
            Public NameSpacex As String
        End Structure

        Private Structure Collection2Type
            Public AddonCnt As Integer
            Public AddonGuid() As String
            Public AddonName() As String
            Public MenuCnt As Integer
            Public Menus() As String
            Public NavigatorCnt As Integer
            Public Navigators() As NavigatorType
        End Structure
        '
        Private CollectionCnt As Integer
        Private Collections() As Collection2Type
        Private Const guidAddonManagerLibraryListCell = "{9767F464-3728-4B7D-904B-3442D7FD03BE}"
        Private Const guidAddonManagerActiveX = "{1DC06F61-1837-419B-AF36-D5CC41E1C9FD}"
        '
        '=====================================================================================
        ' addon api
        '=====================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Me.cp = CP
                returnHtml = getUpload()
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
            End Try
            Return returnHtml
        End Function

        '
        '==========================================================================================================================================
        '   Addon Manager
        '       This is a form that lets you upload an addon
        '       Eventually, this should be substituted with a "Addon Manager Addon" - so the interface can be improved with Contensive recompile
        '==========================================================================================================================================
        '
        Private Function getUpload() As String
            Dim returnResult As String = ""
            Try
                '
                Dim DbUpToDate As Boolean
                Dim GuidFieldName As String
                Dim ErrorMessage As String = ""
                Dim UpgradeOK As Boolean
                Dim installFolder As String
                Dim LibGuids() As String
                Dim FormInput As String
                Dim cnt As Integer
                Dim Ptr As Integer
                Dim PreTableCopy As String
                Dim BodyHTML As String
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim Button As String
                Dim Caption As String
                Dim Description As String
                Dim ButtonList As String
                Dim CollectionFilename As String = ""
                Dim UploadsCnt As Integer
                Dim Doc As New Xml.XmlDocument
                Dim status As String = ""
                Dim collectionsToBeInstalledFromFolder As Boolean
                Dim InstallLibCollectionList As String = ""
                Dim InstallPath As String
                Dim SiteKey As String
                Dim Body As New StringBuilder()
                '
                SiteKey = cp.Site.GetText("sitekey", "")
                If SiteKey = "" Then
                    SiteKey = cp.Utils.CreateGuid()
                    Call cp.Site.SetProperty("sitekey", SiteKey)
                End If
                '
                DbUpToDate = (cp.Site.GetText("buildVersion") >= cp.Version)
                Button = cp.Doc.GetText(RequestNameButton)
                collectionsToBeInstalledFromFolder = False
                GuidFieldName = "ccguid"
                If Button = ButtonCancel Then
                    '
                    ' ----- redirect back to the root
                    '
                    Call cp.Response.Redirect(cp.Site.GetText("adminUrl"))
                Else
                    If Not cp.User.IsAdmin() Then
                        '
                        ' ----- Put up error message
                        '
                        ButtonList = ButtonCancel
                        BodyHTML = cp.Html.p("You must be an administrator to use this tool.")
                    Else
                        '
                        PreTableCopy = "Use this form to upload an add-on collection. If the GUID of the add-on matches one already installed on this server, it will be updated. If the GUID is new, it will be added."
                        installFolder = "CollectionUpload" & cp.Utils.CreateGuid
                        InstallPath = cp.Site.PhysicalFilePath & installFolder & "\"
                        If (Button = ButtonOK) Then
                            '
                            '---------------------------------------------------------------------------------------------
                            ' Reinstall core collection
                            '---------------------------------------------------------------------------------------------
                            '
                            If cp.User.IsDeveloper And cp.Doc.GetBoolean("InstallCore") Then
                                Call cp.Content.DeleteRecords("Add-on Collections", "ccguid='{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}'")
                                UpgradeOK = cp.Addon.installCollectionFromLibrary("{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}", ErrorMessage)
                            End If
                            '
                            '---------------------------------------------------------------------------------------------
                            ' Upload new collection files
                            '---------------------------------------------------------------------------------------------
                            '
                            If cp.Html.ProcessInputFile("MetaFile", cp.privateFiles, installFolder, CollectionFilename) Then
                                Dim taskId As Integer = cp.Utils.installCollectionAsyncFromFile(installFolder & CollectionFilename)
                                status &= "<BR>Uploaded collection file [" & CollectionFilename & "]. Queued for processing as task [" & taskId & "]"
                                UploadsCnt = cp.Doc.GetInteger("UploadCount")
                                For Ptr = 0 To UploadsCnt - 1
                                    If cp.Html.ProcessInputFile("Upload" & Ptr, cp.privateFiles, installFolder, CollectionFilename) Then
                                        taskId = cp.Utils.installCollectionAsyncFromFile(CollectionFilename)
                                        status &= "<BR>Uploaded collection file [" & CollectionFilename & "]. Queued for processing as task [" & taskId & "]"
                                    End If
                                Next
                                status = status & "<BR>Submitted Collection for import."
                            End If
                        End If
                        '
                        ' --------------------------------------------------------------------------------
                        '   Install Library Collections
                        ' --------------------------------------------------------------------------------
                        '
                        If InstallLibCollectionList <> "" Then
                            InstallLibCollectionList = Mid(InstallLibCollectionList, 2)
                            LibGuids = Split(InstallLibCollectionList, ",")
                            cnt = UBound(LibGuids) + 1
                            For Ptr = 0 To cnt - 1
                                cp.Utils.installCollectionAsyncFromLibrary(LibGuids(Ptr))
                            Next
                        End If
                        '
                        ' and delete the install folder if it was created
                        '
                        If cp.privateFiles.folderExists(InstallPath) Then
                            Call cp.privateFiles.deleteFolder(InstallPath)
                        End If
                        '
                        ' --------------------------------------------------------------------------------
                        ' Get Form
                        ' --------------------------------------------------------------------------------
                        '
                        ' --------------------------------------------------------------------------------
                        ' Get the Upload Add-ons tab
                        ' --------------------------------------------------------------------------------
                        '
                        Body = New StringBuilder
                        If Not DbUpToDate Then
                            Call Body.Append("<p>Add-on upload is disabled because your site database needs to be updated.</p>")
                        Else
                            'Call Body.Append(AdminUI.EditTableOpen)

                            'If cp.User.IsDeveloper Then
                            '    Call Body.Append(AdminUI.GetEditRow(cp.Html.CheckBox("InstallCore"), "Reinstall Core Collection", "", False, False, ""))
                            'End If
                            'Call Body.Append(AdminUI.GetEditRow(cp.Html.InputFile("MetaFile"), "Add-on Collection File(s)", "", True, False, ""))
                            FormInput = "" _
                            & "<TABLE id=""UploadInsert"" border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%"">" _
                            & "</Table>" _
                            & "<TABLE border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%"">" _
                            & "<TR><TD align=""left""><a href=""#"" onClick=""InsertUpload(); return false;"">+ Add more files</a></TD></TR>" _
                            & "</Table>" _
                            & cp.Html.Hidden("UploadCount", "1", "UploadCount") _
                            & ""
                            'Call Body.Append(AdminUI.GetEditRow(main, FormInput, "&nbsp;", "", True, False, ""))
                            'Call Body.Append(AdminUI.EditTableClose)
                        End If
                        ' Call main.AddLiveTabEntry("Add&nbsp;Manually", AdminUI.GetEditPanel(main, True, "Install or Update an Add-on Collection.", "Use this form to upload a new or updated Add-on Collection to your site. A collection file can be a single xml configuration file, a single zip file containing the configuration file and other resource files, or a configuration with other resource files uploaded separately. Use the 'Add more files' link to add as many files as you need. When you hit OK, the Collection will be checked, and only submitted if all files are uploaded.", Body.Text), "ccAdminTab")
                        '
                        ' --------------------------------------------------------------------------------
                        ' Build Page from tabs
                        ' --------------------------------------------------------------------------------
                        '
                        ButtonList = ButtonCancel & "," & ButtonOK
                    End If
                    '
                    ' Output the Add-on
                    '
                    Caption = "Add-on Manager"
                    Description = "<div>Use the add-on manager to add and remove Add-ons from your Contensive installation.</div>"
                    If Not DbUpToDate Then
                        Description = Description & "<div style=""Margin-left:50px"">The Add-on Manager is disabled because this site's Database needs to be upgraded.</div>"
                    End If
                    If status <> "" Then
                        Description = Description & "<div style=""Margin-left:50px"">" & status & "</div>"
                    End If
                    ' GetForm_AddonManager = AdminUI.GetBody(main, Caption, ButtonList, "", False, False, Description, "", 0, Content.Text)
                    ' Call main.AddPageTitle("Add-on Manager")
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return returnResult
        End Function
        '        '
        '        '
        '        '
        '        Private Sub GetForm_AddonManager_DeleteNavigatorBranch(EntryName As String, EntryParentID As Integer)
        '            On Error GoTo ErrorTrap
        '            '
        '            Dim cs As CPCSBaseClass = cp.CSNew
        '            Dim EntryID As Integer
        '            '
        '            Call AppendLogFile2(main.ApplicationName, "Deleting Navigator Branch", App.EXEName, "AdminClass", "GetForm_AddonManager_DeleteNavigatorBranch", 0, "", "", False, True, main.ServerLink, "AddonInstall", "")
        '            '
        '            If EntryParentID = 0 Then
        '                cs.open("Navigator Entries", "(name=" & cp.Db.EncodeSQLText(EntryName) & ")and((parentID is null)or(parentid=0))")
        '            Else
        '                cs.open("Navigator Entries", "(name=" & cp.Db.EncodeSQLText(EntryName) & ")and(parentID=" & cp.Db.EncodeSQLNumber(EntryParentID) & ")")
        '            End If
        '            If cs.ok Then
        '                EntryID = cs.getinteger("ID")
        '            End If
        '            Call cs.close()
        '            '
        '            If EntryID <> 0 Then
        '                cs.open("Navigator Entries", "(parentID=" & cp.Db.EncodeSQLNumber(EntryID) & ")")
        '                Do While cs.ok
        '                    Call GetForm_AddonManager_DeleteNavigatorBranch(cs.getText("name"), EntryID)
        '                    Call cs.goNext()
        '                Loop
        '                Call cs.close()
        '                Call cp.Content.Delete("Navigator Entries", EntryID)
        '            End If
        '            '
        '            Exit Sub
        'ErrorTrap:
        '            Call HandleClassTrapError("GetForm_AddonManager_DeleteNavigatorBranch")
        '        End Sub


        '        '
        '        '
        '        '
        '        Private Function GetParentIDFromNameSpace(ContentName As String, NameSpacex As String) As Integer
        '            On Error GoTo ErrorTrap
        '            '
        '            Dim NameSplit() As String
        '            Dim ParentNameSpace As String
        '            Dim ParentName As String
        '            Dim ParentID As Integer
        '            Dim Pos As Integer
        '            Dim cs As CPCSBaseClass = cp.CSNew
        '            '
        '            GetParentIDFromNameSpace = 0
        '            If NameSpacex <> "" Then
        '                'ParentName = ParentNameSpace
        '                Pos = InStr(1, NameSpacex, ".")
        '                If Pos = 0 Then
        '                    ParentName = NameSpacex
        '                    ParentNameSpace = ""
        '                Else
        '                    ParentName = Mid(NameSpacex, Pos + 1)
        '                    ParentNameSpace = Mid(NameSpacex, 1, Pos - 1)
        '                End If
        '                If ParentNameSpace = "" Then
        '                    cs.open(ContentName, "(name=" & cp.Db.EncodeSQLText(ParentName) & ")and((parentid is null)or(parentid=0))", "ID", , , , "ID")
        '                    If cs.ok Then
        '                        GetParentIDFromNameSpace = cs.getinteger("ID")
        '                    End If
        '                    Call cs.close()
        '                Else
        '                    ParentID = GetParentIDFromNameSpace(ContentName, ParentNameSpace)
        '                    cs.open(ContentName, "(name=" & cp.Db.EncodeSQLText(ParentName) & ")and(parentid=" & ParentID & ")", "ID", , , , "ID")
        '                    If cs.ok Then
        '                        GetParentIDFromNameSpace = cs.getinteger("ID")
        '                    End If
        '                    Call cs.close()
        '                End If
        '            End If
        '            '
        '            Exit Function
        '            '
        '            ' ----- Error Trap
        '            '
        'ErrorTrap:
        '            Call HandleClassTrapError("GetParentIDFromNameSpace", False)
        '        End Function
        '        '
        '        '
        '        '
        '        Private Function getLayout(layoutGuid As String, DefaultName As String) As String
        '            On Error GoTo ErrorTrap
        '            '
        '            Dim cs As CPCSBaseClass = cp.CSNew
        '            '
        '            cs.open("layouts", "ccguid=" & cp.Db.EncodeSQLText(layoutGuid))
        '            If Not cs.ok Then
        '                Call cs.close()
        '                cs = main.InsertCSContent("layouts")
        '                If cs.ok Then
        '                    Call cs.SetField("ccguid", layoutGuid)
        '                    Call cs.SetField("name", DefaultName)
        '                    Call cs.SetField("layout", "<!-- layout record created " & Now & " -->")
        '                End If
        '            End If
        '            If cs.ok Then
        '                getLayout = cs.get(cs, "layout")
        '            End If
        '            Call cs.close()
        '            '
        '            Exit Function
        '            '
        '            ' ----- Error Trap
        '            '
        'ErrorTrap:
        '            Call HandleClassTrapError("getLayout", False)
        '        End Function
        '        '
        '        '========================================================================
        '        '   HandleError
        '        '       Logs the error and either resumes next, or raises it to the next level
        '        '========================================================================
        '        '
        '        Public Sub HandleError(cp As CPBaseClass, ByVal ex As Exception, ByVal className As String, ByVal methodName As String, ByVal cause As String)
        '            Try
        '                Dim errMsg As String = className & "." & methodName & ", cause=[" & cause & "], ex=[" & ex.ToString & "]"
        '                cp.Site.ErrorReport(ex, errMsg)
        '            Catch exIgnore As Exception
        '                '
        '            End Try
        '        End Sub
        '        '
        '        Public Sub HandleError(cp As CPBaseClass, ByVal ex As Exception)
        '            Dim frame As StackFrame = New StackFrame(1)
        '            Dim method As System.Reflection.MethodBase = frame.GetMethod()
        '            Dim type As System.Type = method.DeclaringType()
        '            Dim methodName As String = method.Name
        '            '
        '            Call HandleError(cp, ex, type.Name, methodName, "unexpected exception")
        '        End Sub
    End Class
End Namespace
