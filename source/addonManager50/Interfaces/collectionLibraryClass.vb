
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports System.Xml
Imports adminFramework

Namespace Contensive.Addons.AddonManager
    '
    ' Sample Vb addon
    '
    Public Class collectionLibraryClass
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
        ''' <summary>
        ''' addon api
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Me.cp = CP
                returnHtml = GetCollectionLibrary()
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
        Private Function GetCollectionLibrary() As String
            Dim returnResult As String = ""
            Try
                Dim CollectionFilename As String = ""
                Dim Doc As New XmlDocument
                Dim SiteKey As String
                '
                SiteKey = cp.Site.GetText("sitekey", "")
                If SiteKey = "" Then
                    SiteKey = cp.Utils.CreateGuid()
                    Call cp.Site.SetProperty("sitekey", SiteKey)
                End If
                'Dim DisplaySystem As Boolean
                '
                Dim DbUpToDate As Boolean = (cp.Site.GetText("buildVersion") >= cp.Version)
                Dim Button As String = cp.Doc.GetText(RequestNameButton)
                Dim collectionsToBeInstalledFromFolder As Boolean = False
                Dim GuidFieldName As String = "ccguid"
                If Button = ButtonCancel Then
                    '
                    ' ----- redirect back to the root
                    '
                    Call cp.Response.Redirect(cp.Site.GetText("adminUrl"))
                Else
                    '
                    Dim form As New formSimpleClass
                    Dim BodyHTML As String = ""
                    Dim status As String = ""
                    If Not cp.User.IsAdmin() Then
                        '
                        ' ----- Put up error message
                        '
                        Dim ButtonList As String = ButtonCancel
                        BodyHTML = cp.Html.p("You must be an administrator to use this tool.")
                    Else
                        '
                        Dim PreTableCopy As String = "Use this form to upload an add-on collection. If the GUID of the add-on matches one already installed on this server, it will be updated. If the GUID is new, it will be added."
                        Dim installFolder As String = "CollectionUpload" & cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "")
                        Dim InstallPath As String = installFolder & "\"
                        Dim CollectionGUID As String = ""
                        If (Button = ButtonOK) Then
                            '
                            '---------------------------------------------------------------------------------------------
                            ' Download and install Collections from the Collection Library
                            '---------------------------------------------------------------------------------------------
                            '
                            Dim InstallLibCollectionList As String = ""
                            Dim Ptr As Integer
                            If cp.Doc.GetText("LibraryRow") <> "" Then
                                Ptr = cp.Doc.GetInteger("LibraryRow")
                                CollectionGUID = cp.Doc.GetText("LibraryRowguid" & Ptr)
                                InstallLibCollectionList = InstallLibCollectionList & "," & CollectionGUID
                            End If
                            '
                            If InstallLibCollectionList <> "" Then
                                InstallLibCollectionList = Mid(InstallLibCollectionList, 2)
                                Dim LibGuids() As String = Split(InstallLibCollectionList, ",")
                                Dim cnt As Integer = UBound(LibGuids) + 1
                                For Ptr = 0 To cnt - 1
                                    Dim UpgradeOK As Boolean
                                    If (cp.Version < "5") Then
                                        '
                                        ' -- use [addonManager-InstallFromLibraryC41]
                                        cp.Doc.SetProperty("collectionGuid", LibGuids(Ptr))
                                        UpgradeOK = cp.Utils.EncodeBoolean(cp.Utils.ExecuteAddon(legacyMethodInstallFromLibraryC41))
                                    Else
                                        Dim ErrorMessage As String = ""
                                        '
                                        ' -- use v5 method
                                        UpgradeOK = v5InstallController.installCollectionFromLibrary(cp, LibGuids(Ptr), ErrorMessage)
                                    End If
                                Next
                            End If
                        End If
                        '
                        ' --------------------------------------------------------------------------------
                        ' Get Form
                        ' --------------------------------------------------------------------------------
                        ' Get the Collection Library tab
                        ' --------------------------------------------------------------------------------
                        '
                        Dim ColumnCnt As Integer = 4
                        Dim PageNumber As Integer = 1
                        Dim ColCaption() As String
                        ReDim ColCaption(3)
                        Dim ColAlign() As String
                        ReDim ColAlign(3)
                        Dim ColWidth() As String
                        ReDim ColWidth(3)
                        Dim ColSortable() As Boolean
                        ReDim ColSortable(3)
                        'ReDim Cells3(10, 4)
                        '
                        ColCaption(0) = "Install"
                        ColAlign(0) = "center"
                        ColWidth(0) = "50"
                        ColSortable(0) = False
                        '
                        ColCaption(1) = "Name"
                        ColAlign(1) = "left"
                        ColWidth(1) = "200"
                        ColSortable(1) = False
                        '
                        ColCaption(2) = "Last&nbsp;Updated"
                        ColAlign(2) = "right"
                        ColWidth(2) = "200"
                        ColSortable(2) = False
                        '
                        ColCaption(3) = "Description"
                        ColAlign(3) = "left"
                        ColWidth(3) = "99%"
                        ColSortable(3) = False
                        '
                        Dim LibCollections As XmlDocument = New Xml.XmlDocument
                        Call LibCollections.Load("http://support.contensive.com/GetCollectionList?iv=" & cp.Version & "&key=" & cp.Utils.EncodeRequestVariable(SiteKey) & "&name=" & cp.Utils.EncodeRequestVariable(cp.Site.Name) & "&primaryDomain=" & cp.Utils.EncodeRequestVariable(cp.Site.DomainPrimary))
                        If True Then
                            Dim RowPtr As Integer
                            If (LibCollections.DocumentElement.Name.ToLower() <> CollectionListRootNode.ToLower()) Then
                                Dim UserError As String = "There was an error reading the Collection Library file. The '" & CollectionListRootNode & "' element was not found."
                                status = status & "<BR>" & UserError
                                cp.UserError.Add(UserError)
                            Else
                                '
                                ' Go through file to validate the XML, and build status message -- since service process can not communicate to user
                                '
                                RowPtr = 0
                                'Content = ""
                                Dim cellTemplate As String = My.Resources.AddonManagerLibraryListCell
                                For Each CDef_Node As Xml.XmlNode In LibCollections.DocumentElement.ChildNodes
                                    Dim Cell As String = cellTemplate
                                    Dim CollectionImageLink As String = ""
                                    Dim CollectionCheckbox As String = ""
                                    Dim CollectionName As String = ""
                                    Dim CollectionModifiedDate As Date = Date.MinValue
                                    Dim CollectionModifiedDateCaption As String = ""
                                    Dim CollectionDescription As String = ""
                                    Dim CollectionHelpLink As String = ""
                                    Dim CollectionDemoLink As String = ""
                                    Select Case LCase(CDef_Node.Name)
                                        Case "collection"
                                            Dim CollectionContensiveVersion As String = ""
                                            Dim CollectionNode As XmlNode
                                            '
                                            ' Read the collection
                                            '
                                            For Each CollectionNode In CDef_Node.ChildNodes
                                                Select Case CollectionNode.Name.ToLower()
                                                    Case "name"
                                                        '
                                                        ' Name
                                                        '
                                                        CollectionName = CollectionNode.InnerText
                                                    Case "helplink"
                                                        '
                                                        ' helpLink
                                                        '
                                                        CollectionHelpLink = CollectionNode.InnerText
                                                    Case "demolink"
                                                        '
                                                        ' demoLink
                                                        '
                                                        CollectionDemoLink = CollectionNode.InnerText
                                                    Case "guid"
                                                        '
                                                        ' Guid
                                                        '
                                                        CollectionGUID = CollectionNode.InnerText
                                                    Case "version"
                                                        '
                                                        ' Version
                                                        '
                                                        Dim CollectionVersion As String = CollectionNode.InnerText
                                                    Case "description"
                                                        '
                                                        ' Version
                                                        '
                                                        CollectionDescription = CollectionNode.InnerText
                                                    Case "imagelink"
                                                        '
                                                        ' Version
                                                        '
                                                        CollectionImageLink = CollectionNode.InnerText
                                                    Case "contensiveversion"
                                                        '
                                                        ' legacy, used by 4.1 to auto-upgrade
                                                        '
                                                        CollectionContensiveVersion = CollectionNode.InnerText
                                                    Case "lastmodifieddate"
                                                        '
                                                        ' Version
                                                        '
                                                        CollectionModifiedDate = Date.MinValue
                                                        If IsDate(CollectionNode.InnerText) Then
                                                            CollectionModifiedDate = CDate(CollectionNode.InnerText)
                                                        End If
                                                        If (CollectionModifiedDate <= Date.MinValue) Then
                                                            CollectionModifiedDateCaption = "unknown"
                                                        Else
                                                            CollectionModifiedDateCaption = CollectionModifiedDate.Date.ToShortDateString
                                                        End If
                                                End Select
                                            Next
                                            If CollectionImageLink = "" Then
                                                CollectionImageLink = "/addonManager/libraryNoImage.jpg"
                                            End If
                                            If CollectionModifiedDateCaption = "" Then
                                                CollectionModifiedDateCaption = "unknown"
                                            End If
                                            If CollectionDescription = "" Then
                                                CollectionDescription = "No description is available for this add-on collection."
                                            End If
                                            Dim showAddon As Boolean = False
                                            If CollectionName = "" Then
                                            Else
                                                If CollectionGUID = "" Then
                                                Else
                                                    Dim cs As CPCSBaseClass = cp.CSNew
                                                    Dim modifiedDate As Date = Date.MinValue
                                                    Dim IsOnSite As Boolean = False
                                                    If cs.Open("Add-on Collections", GuidFieldName & "=" & cp.Db.EncodeSQLText(CollectionGUID), "", True, "ID,ModifiedDate", 9990, 1) Then
                                                        modifiedDate = cs.GetDate("ModifiedDate")
                                                        IsOnSite = True
                                                    End If
                                                    Call cs.Close()
                                                    CollectionCheckbox = "<input TYPE=""CheckBox"" NAME=""LibraryRow"" VALUE=""" & RowPtr & """ onClick=""clearLibraryRows('" & RowPtr & "');"">" & cp.Html.Hidden("LibraryRowGuid" & RowPtr, CollectionGUID) & cp.Html.Hidden("LibraryRowName" & RowPtr, CollectionName)
                                                    If (Not IsOnSite) Then
                                                        '
                                                        ' -- not installed
                                                        showAddon = True
                                                        CollectionCheckbox &= "&nbsp;Install"
                                                    ElseIf (modifiedDate >= CollectionModifiedDate) Then
                                                        '
                                                        ' -- up to date, reinstall
                                                        showAddon = True
                                                        CollectionCheckbox &= "&nbsp;Reinstall"
                                                    Else
                                                        '
                                                        ' -- old version, upgrade
                                                        showAddon = True
                                                        CollectionCheckbox &= "&nbsp;Upgrade"
                                                    End If
                                                End If
                                            End If
                                            If CollectionDemoLink <> "" Then
                                                CollectionDescription = CollectionDescription & "<div class=""amDemoLink""><a target=""_blank"" href=""" & CollectionDemoLink & """>Demo</a></div>"
                                            End If
                                            If CollectionHelpLink <> "" Then
                                                CollectionDescription = CollectionDescription & "<div class=""amHelpLink""><a target=""_blank"" href=""" & CollectionHelpLink & """>Reference</a></div>"
                                            End If
                                            If showAddon Then
                                                Cell = Replace(Cell, "##imageLink##", CollectionImageLink)
                                                Cell = Replace(Cell, "##checkbox##", CollectionCheckbox)
                                                Cell = Replace(Cell, "##name##", CollectionName)
                                                Cell = Replace(Cell, "##date##", CollectionModifiedDateCaption)
                                                Cell = Replace(Cell, "##description##", CollectionDescription)
                                                BodyHTML = BodyHTML & Cell
                                            End If
                                            RowPtr = RowPtr + 1
                                    End Select
                                Next
                            End If
                            BodyHTML = "" _
                            & cr & "<script language=""JavaScript"">" _
                            & "function clearLibraryRows(r) {" _
                            & "var c,p;" _
                            & "c=document.getElementsByName('LibraryRow');" _
                                & "for (p=0;p<c.length;p++){" _
                                    & "if(c[p].value!=r)c[p].checked=false;" _
                                & "}" _
                            & "" _
                            & "}" _
                            & "</script>" _
                            & "<input type=hidden name=LibraryCnt value=""" & RowPtr & """>" _
                            & cr & "<div style=""width:100%"">" _
                            & kmaIndent(BodyHTML) _
                            & cr & "</div>" _
                            & ""
                            form.body = BodyHTML
                            form.description = "Select Add-ons to install from the Contensive Add-on Library. Please select only one at a time. Click OK to install the selected Add-on. The site may need to be stopped during the installation, but will be available again in approximately one minute"
                        End If
                        '
                        ' --------------------------------------------------------------------------------
                        ' Build Page from tabs
                        ' --------------------------------------------------------------------------------
                        '
                        form.addFormButton(ButtonCancel)
                        form.addFormButton(ButtonOK)
                    End If
                    '
                    ' Output the Add-on
                    form.title = "Add-on Library"
                    form.description = ""
                    If Not DbUpToDate Then
                        form.description = form.description & "<div style=""Margin-left:50px"">The Add-on Manager is disabled because this site's Database needs to be upgraded.</div>"
                    End If
                    If status <> "" Then
                        form.description = form.description & "<div style=""Margin-left:50px"">" & status & "</div>"
                    End If
                    returnResult = form.getHtml(cp)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return returnResult
        End Function
    End Class
End Namespace
