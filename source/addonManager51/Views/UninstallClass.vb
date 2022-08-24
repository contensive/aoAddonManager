
Imports Contensive.Addons.PortalFramework
Imports Contensive.BaseClasses
Imports Contensive.Models.Db

Namespace Contensive.Addons.AddonManager51
    '
    Public Class UninstallClass
        Inherits AddonBaseClass
        '
        ' injected objects -- do not dispose
        '
        Private cp As CPBaseClass
        ' 
        ' class scope
        '
        Private Const RequestNameButton As String = "button"
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
        '=====================================================================================
        ''' <summary>
        ''' addon api
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Try
                Me.cp = CP
                Return getUnistall()
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
        '       
        '==========================================================================================================================================
        ''' <summary>
        ''' Addon Manager,  This is a form that lets you upload an addon
        ''' Eventually, this should be substituted with a "Addon Manager Addon" - so the interface can be improved with Contensive recompile
        ''' </summary>
        ''' <returns></returns>
        Private Function getUnistall() As String
            Dim returnResult As String = ""
            Try

                Dim DisplaySystem As Boolean
                Dim DbUpToDate As Boolean
                Dim GuidFieldName As String
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
                Dim SiteKey As String
                Dim form As New PortalFramework.ReportListClass(cp)
                '
                SiteKey = cp.Site.GetText("sitekey", "")
                If String.IsNullOrEmpty(SiteKey) Then
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
                                                Call getForm_AddonManager_DeleteNavigatorBranch(cp, TargetCollectionName, AddonNavigatorID)
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
                        form.columnCaptionClass = AfwStyles.afwTextAlignCenter + " " + AfwStyles.afwWidth50px
                        form.columnCellClass = AfwStyles.afwTextAlignCenter
                        form.columnDownloadable = False
                        form.columnName = ""
                        form.columnSortable = False
                        form.columnVisible = True
                        '
                        form.addColumn()
                        form.columnCaption = "Name"
                        form.columnCaptionClass = AfwStyles.afwTextAlignLeft
                        form.columnCellClass = AfwStyles.afwTextAlignLeft
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
                    If Not String.IsNullOrEmpty(status) Then
                        form.description &= "<div style=""Margin-left:50px"">" & status & "</div>"
                    End If
                    returnResult = form.getHtml(cp)
                End If
                Return returnResult
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
    End Class
End Namespace
