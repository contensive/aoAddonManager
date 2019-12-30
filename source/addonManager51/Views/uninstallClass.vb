
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Contensive.Addons.AddonManager51
    '
    ' Sample Vb addon
    '
    Public Class uninstallClass
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
        Private Const ButtonOK As String = " OK "
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
                returnHtml = getUnistall()
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
        Private Function getUnistall() As String
            Dim returnResult As String = ""
            Try

                Dim DisplaySystem As Boolean
                Dim DbUpToDate As Boolean
                Dim GuidFieldName As String
                Dim installFolder As String
                Dim AddonNavigatorID As Integer
                Dim TargetCollectionName As String
                Dim addonid As Integer
                Dim cnt As Integer
                Dim Ptr As Integer
                Dim BodyHTML As String
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim Button As String
                Dim ButtonList As String
                Dim CollectionFilename As String = ""
                Dim Doc As New Xml.XmlDocument
                Dim status As String = ""
                Dim collectionsToBeInstalledFromFolder As Boolean
                Dim TargetCollectionID As Integer
                Dim InstallPath As String
                Dim SiteKey As String
                Dim form As New adminFramework.ReportListClass(cp)
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
                        'installFolder = "CollectionUpload" & cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "")
                        'InstallPath = cp.Site.PhysicalFilePath & installFolder & "\"
                        If (Button = ButtonOK) Then
                            '
                            '---------------------------------------------------------------------------------------------
                            ' Delete collections
                            '   Before deleting each addon, make sure it is not in another collection
                            '---------------------------------------------------------------------------------------------
                            '
                            cnt = cp.Doc.GetInteger("accnt")
                            If cnt > 0 Then
                                For Ptr = 0 To cnt - 1
                                    If cp.Doc.GetBoolean("ac" & Ptr) Then
                                        TargetCollectionID = cp.Doc.GetInteger("acID" & Ptr)
                                        TargetCollectionName = cp.Content.GetRecordName("Add-on Collections", TargetCollectionID)
                                        '
                                        ' Clean up rules associating this collection to other objects
                                        '
                                        Call cp.Content.Delete("Add-on Collection CDef Rules", "collectionid=" & TargetCollectionID)
                                        Call cp.Content.Delete("Add-on Collection Module Rules", "collectionid=" & TargetCollectionID)
                                        '
                                        ' Delete any addons from this collection
                                        '
                                        If cs.Open("add-ons", "collectionid=" & TargetCollectionID) Then
                                            Do
                                                '
                                                ' Clean up the rules that might have pointed to the addon
                                                '
                                                addonid = cs.GetInteger("id")
                                                Call cp.Content.Delete("Admin Menuing", "addonid=" & addonid)
                                                Call cp.Content.Delete("Shared Styles Add-on Rules", "addonid=" & addonid)
                                                Call cp.Content.Delete("Add-on Scripting Module Rules", "addonid=" & addonid)
                                                Call cp.Content.Delete("Add-on Include Rules", "addonid=" & addonid)
                                                Call cp.Content.Delete("Add-on Include Rules", "includedaddonid=" & addonid)
                                                cs.GoNext()
                                            Loop While cs.OK
                                        End If
                                        cs.Close()
                                        Call cp.Content.Delete("add-ons", "collectionid=" & TargetCollectionID)
                                        '
                                        ' Delete the navigator entry for the collection under 'Add-ons'
                                        '
                                        If TargetCollectionID > 0 Then
                                            AddonNavigatorID = 0
                                            cs.Open("Navigator Entries", "name='Manage Add-ons' and ((parentid=0)or(parentid is null))")
                                            If cs.OK Then
                                                AddonNavigatorID = cs.GetInteger("ID")
                                            End If
                                            Call cs.Close()
                                            If AddonNavigatorID > 0 Then
                                                Call GetForm_AddonManager_DeleteNavigatorBranch(cp, TargetCollectionName, AddonNavigatorID)
                                            End If
                                            '
                                            ' Now delete the Collection record
                                            '
                                            cp.Content.Delete("Add-on Collections", "id=" & TargetCollectionID)
                                            '
                                            ' Delete Navigator Entries set as installed by the collection (this may be all that is needed)
                                            '
                                            Call cp.Content.Delete("Navigator Entries", "installedbycollectionid=" & TargetCollectionID)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        '
                        ' --------------------------------------------------------------------------------
                        ' Current Collections Tab
                        ' --------------------------------------------------------------------------------
                        '
                        form.addColumn()
                        form.columnCaption = "Del"
                        form.columnCaptionClass = adminFramework.AfwStyles.afwTextAlignCenter + " " + adminFramework.AfwStyles.afwWidth50px
                        form.columnCellClass = adminFramework.AfwStyles.afwTextAlignCenter
                        form.columnDownloadable = False
                        form.columnName = ""
                        form.columnSortable = False
                        form.columnVisible = True
                        '
                        form.addColumn()
                        form.columnCaption = "Name"
                        form.columnCaptionClass = adminFramework.AfwStyles.afwTextAlignLeft
                        form.columnCellClass = adminFramework.AfwStyles.afwTextAlignLeft
                        form.columnDownloadable = False
                        form.columnName = ""
                        form.columnSortable = False
                        form.columnVisible = True
                        '
                        DisplaySystem = False
                        Dim addonCollectionList As List(Of AddonCollectionModel) = AddonCollectionModel.createList(Of AddonCollectionModel)(cp, "", "name")
                        Dim rowPtr As Integer = 0
                        For Each item In addonCollectionList
                            form.addRow()
                            form.setCell(cp.Html.CheckBox("AC" & rowPtr) & cp.Html.Hidden("ACID" & rowPtr, item.id.ToString()))
                            form.setCell(item.name)
                            rowPtr += 1
                        Next
                        form.addFormButton(ButtonCancel)
                        form.addFormButton(ButtonOK)
                        form.addFormHidden("accnt", addonCollectionList.Count.ToString())
                    End If
                    '
                    ' Output the Add-on
                    '
                    form.name = "Uninstall Collections"
                    form.description = "To remove collections, select them from the list and click the Uninstall button."
                    If Not DbUpToDate Then
                        form.description &= "<div style=""Margin-left:50px"">Warning: The site's Database needs to be upgraded.</div>"
                    End If
                    If status <> "" Then
                        form.description &= "<div style=""Margin-left:50px"">" & status & "</div>"
                    End If
                    returnResult = form.getHtml(cp)
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
