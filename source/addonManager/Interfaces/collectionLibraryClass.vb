
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports System.Xml
Imports adminFramework

Namespace Contensive.Addons.AddonManager51
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
        ' addon api
        '=====================================================================================
        '
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
                '
                Dim form As New formSimpleClass
                Dim showAddon As Boolean
                'Dim cellTemplate As String
                'Dim Cell As String
                'Dim CollectionCheckbox As String
                'Dim IsOnSite As Boolean
                'Dim DisplaySystem As Boolean
                Dim DbUpToDate As Boolean
                Dim GuidFieldName As String
                'Dim DateValue As Date
                Dim ErrorMessage As String = ""
                Dim UpgradeOK As Boolean
                Dim LibCollections As Xml.XmlDocument
                Dim installFolder As String
                Dim LibGuids() As String
                'Dim CollectionName As String
                Dim CollectionGUID As String = ""
                Dim CollectionVersion As String
                'Dim CollectionDescription As String
                'Dim CollectionImageLink As String
                Dim CollectionContensiveVersion As String = ""
                'Dim CollectionLastChangeDate As String
                Dim cnt As Integer
                Dim Ptr As Integer
                Dim RowPtr As Integer
                Dim PageNumber As Integer
                Dim ColumnCnt As Integer
                Dim ColCaption() As String
                Dim ColAlign() As String
                Dim ColWidth() As String
                Dim ColSortable() As Boolean
                Dim PreTableCopy As String
                Dim BodyHTML As String = ""
                'Dim cs As CPCSBaseClass = cp.CSNew
                Dim UserError As String
                Dim Button As String
                Dim ButtonList As String
                Dim CollectionFilename As String = ""
                Dim Doc As New Xml.XmlDocument
                Dim CollectionNode As Xml.XmlNode
                Dim status As String = ""
                Dim collectionsToBeInstalledFromFolder As Boolean
                Dim InstallLibCollectionList As String = ""
                Dim InstallPath As String
                Dim SiteKey As String
                'Dim CollectionHelpLink As String
                'Dim CollectionDemoLink As String
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
                        installFolder = "CollectionUpload" & cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "")
                        InstallPath = installFolder & "\"
                        If (Button = ButtonOK) Then
                            '
                            '---------------------------------------------------------------------------------------------
                            ' Download and install Collections from the Collection Library
                            '---------------------------------------------------------------------------------------------
                            '
                            InstallLibCollectionList = ""
                            If cp.Doc.GetText("LibraryRow") <> "" Then
                                Ptr = cp.Doc.GetInteger("LibraryRow")
                                CollectionGUID = cp.Doc.GetText("LibraryRowguid" & Ptr)
                                InstallLibCollectionList = InstallLibCollectionList & "," & CollectionGUID
                            End If
                            '
                            If InstallLibCollectionList <> "" Then
                                InstallLibCollectionList = Mid(InstallLibCollectionList, 2)
                                LibGuids = Split(InstallLibCollectionList, ",")
                                cnt = UBound(LibGuids) + 1
                                For Ptr = 0 To cnt - 1
                                    '
                                    ' -- use v5 method
                                    UpgradeOK = v5InstallController.installCollectionFromLibrary(cp, LibGuids(Ptr), ErrorMessage)
                                Next
                            End If


                            ''
                            ''---------------------------------------------------------------------------------------------
                            '' Delete collections
                            ''   Before deleting each addon, make sure it is not in another collection
                            ''---------------------------------------------------------------------------------------------
                            ''
                            'cnt = cp.Doc.GetInteger("accnt")
                            'If cnt > 0 Then
                            '    For Ptr = 0 To cnt - 1
                            '        If cp.Doc.GetBoolean("ac" & Ptr) Then
                            '            TargetCollectionID = cp.Doc.GetInteger("acID" & Ptr)
                            '            TargetCollectionName = cp.Content.GetRecordName("Add-on Collections", TargetCollectionID)
                            '            '
                            '            ' Clean up rules associating this collection to other objects
                            '            '
                            '            Call cp.Content.Delete("Add-on Collection CDef Rules", "collectionid=" & TargetCollectionID)
                            '            Call cp.Content.Delete("Add-on Collection Module Rules", "collectionid=" & TargetCollectionID)
                            '            '
                            '            ' Delete any addons from this collection
                            '            '
                            '            If cs.Open("add-ons", "collectionid=" & TargetCollectionID) Then
                            '                Do
                            '                    '
                            '                    ' Clean up the rules that might have pointed to the addon
                            '                    '
                            '                    addonid = cs.GetInteger("id")
                            '                    Call cp.Content.Delete("Admin Menuing", "addonid=" & addonid)
                            '                    Call cp.Content.Delete("Shared Styles Add-on Rules", "addonid=" & addonid)
                            '                    Call cp.Content.Delete("Add-on Scripting Module Rules", "addonid=" & addonid)
                            '                    Call cp.Content.Delete("Add-on Include Rules", "addonid=" & addonid)
                            '                    Call cp.Content.Delete("Add-on Include Rules", "includedaddonid=" & addonid)
                            '                    cs.GoNext()
                            '                Loop While cs.OK
                            '            End If
                            '            cs.Close()
                            '            Call cp.Content.Delete("add-ons", "collectionid=" & TargetCollectionID)
                            '            '
                            '            ' Delete the navigator entry for the collection under 'Add-ons'
                            '            '
                            '            If TargetCollectionID > 0 Then
                            '                AddonNavigatorID = 0
                            '                cs.Open("Navigator Entries", "name='Manage Add-ons' and ((parentid=0)or(parentid is null))")
                            '                If cs.OK Then
                            '                    AddonNavigatorID = cs.GetInteger("ID")
                            '                End If
                            '                Call cs.Close()
                            '                If AddonNavigatorID > 0 Then
                            '                    Call GetForm_AddonManager_DeleteNavigatorBranch(cp, TargetCollectionName, AddonNavigatorID)
                            '                End If
                            '                '
                            '                ' Now delete the Collection record
                            '                '
                            '                cp.Content.Delete("Add-on Collections", "id=" & TargetCollectionID)
                            '                '
                            '                ' Delete Navigator Entries set as installed by the collection (this may be all that is needed)
                            '                '
                            '                Call cp.Content.Delete("Navigator Entries", "installedbycollectionid=" & TargetCollectionID)
                            '            End If
                            '        End If
                            '    Next
                            'End If
                            ''
                            ''---------------------------------------------------------------------------------------------
                            '' Delete Add-on Collections
                            ''---------------------------------------------------------------------------------------------
                            ''
                            'cnt = cp.Doc.GetInteger("aocnt")
                            'If cnt > 0 Then
                            '    For Ptr = 0 To cnt - 1
                            '        If cp.Doc.GetBoolean("ao" & Ptr) Then
                            '            cp.Content.Delete("Add-on Collections", "id=" & cp.Doc.GetInteger("aoID" & Ptr))
                            '        End If
                            '    Next
                            'End If
                            ''
                            ''---------------------------------------------------------------------------------------------
                            '' Reinstall core collection
                            ''---------------------------------------------------------------------------------------------
                            ''
                            'If cp.User.IsDeveloper And cp.Doc.GetBoolean("InstallCore") Then
                            '    Call cp.Content.Delete("Add-on Collections", "ccguid='{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}'")
                            '    UpgradeOK = cp.Addon.installCollectionFromLibrary("{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}", ErrorMessage)
                            'End If
                            ''
                            ''---------------------------------------------------------------------------------------------
                            '' Upload new collection files
                            ''---------------------------------------------------------------------------------------------
                            ''
                            'Dim ignoreTaskId As Integer
                            'cp.privateFiles.saveUpload("metafile", installFolder, CollectionFilename)
                            'If (cp.privateFiles.saveUpload("metafile", installFolder, CollectionFilename)) Then
                            '    ignoreTaskId = cp.Utils.installCollectionFromFile(installFolder & CollectionFilename)
                            '    status &= "<BR>Uploaded collection file [" & CollectionFilename & "]. Queued for processing as task [" & ignoreTaskId & "]"
                            'End If
                            'UploadsCnt = cp.Doc.GetInteger("UploadCount")
                            'For Ptr = 0 To UploadsCnt - 1
                            '    If (cp.privateFiles.saveUpload("Upload" & Ptr, installFolder, CollectionFilename)) Then
                            '        ignoreTaskId = cp.Utils.installCollectionFromFile(CollectionFilename)
                            '        status &= "<BR>Uploaded collection file [" & CollectionFilename & "]. Queued for processing as task [" & ignoreTaskId & "]"
                            '    End If
                            'Next
                        End If
                        ''
                        '' --------------------------------------------------------------------------------
                        ''   Install Library Collections
                        '' --------------------------------------------------------------------------------
                        ''
                        'If InstallLibCollectionList <> "" Then
                        '    InstallLibCollectionList = Mid(InstallLibCollectionList, 2)
                        '    LibGuids = Split(InstallLibCollectionList, ",")
                        '    cnt = UBound(LibGuids) + 1
                        '    For Ptr = 0 To cnt - 1
                        '        cp.Utils.installCollectionFromLibrary(LibGuids(Ptr))
                        '    Next
                        'End If
                        ''
                        '' and delete the install folder if it was created
                        ''
                        'If cp.privateFiles.folderExists(InstallPath) Then
                        '    Call cp.privateFiles.deleteFolder(InstallPath)
                        'End If
                        '
                        ' --------------------------------------------------------------------------------
                        ' Get Form
                        ' --------------------------------------------------------------------------------
                        ' Get the Collection Library tab
                        ' --------------------------------------------------------------------------------
                        '
                        ColumnCnt = 4
                        PageNumber = 1
                        ReDim ColCaption(3)
                        ReDim ColAlign(3)
                        ReDim ColWidth(3)
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
                        LibCollections = New Xml.XmlDocument
                        Call LibCollections.Load("http://support.contensive.com/GetCollectionList?iv=" & cp.Version & "&key=" & cp.Utils.EncodeRequestVariable(SiteKey) & "&name=" & cp.Utils.EncodeRequestVariable(cp.Site.Name) & "&primaryDomain=" & cp.Utils.EncodeRequestVariable(cp.Site.DomainPrimary))
                        If True Then
                            If (LibCollections.DocumentElement.Name.ToLower() <> CollectionListRootNode.ToLower()) Then
                                UserError = "There was an error reading the Collection Library file. The '" & CollectionListRootNode & "' element was not found."
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
                                    Dim CollectionLastChangeDate As Date = Date.MinValue
                                    Dim CollectionLastChangeDateCaption As String = ""
                                    Dim CollectionDescription As String = ""
                                    Dim CollectionHelpLink As String = ""
                                    Dim CollectionDemoLink As String = ""
                                    Select Case LCase(CDef_Node.Name)
                                        Case "collection"
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
                                                        CollectionVersion = CollectionNode.InnerText
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
                                                        ' Version
                                                        '
                                                        CollectionContensiveVersion = CollectionNode.InnerText
                                                    Case "lastchangedate"
                                                        '
                                                        ' Version
                                                        '
                                                        CollectionLastChangeDate = Date.MinValue
                                                        If IsDate(CollectionNode.InnerText) Then
                                                            CollectionLastChangeDate = CDate(CollectionNode.InnerText)
                                                        End If
                                                        If (CollectionLastChangeDate <= Date.MinValue) Then
                                                            CollectionLastChangeDateCaption = "unknown"
                                                        Else
                                                            CollectionLastChangeDateCaption = CollectionLastChangeDate.Date.ToShortDateString
                                                        End If
                                                End Select
                                            Next
                                            If CollectionImageLink = "" Then
                                                CollectionImageLink = "/addonManager/libraryNoImage.jpg"
                                            End If
                                            If CollectionLastChangeDateCaption = "" Then
                                                CollectionLastChangeDateCaption = "unknown"
                                            End If
                                            If CollectionDescription = "" Then
                                                CollectionDescription = "No description is available for this add-on collection."
                                            End If
                                            'If RowPtr >= UBound(Cells3, 1) Then
                                            '    ReDim Preserve Cells3(RowPtr + 100, ColumnCnt)
                                            'End If
                                            showAddon = False
                                            If CollectionName = "" Then
                                                'Cells3(RowPtr, 0) = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""0"" disabled>"
                                                'Cells3(RowPtr, 1) = "no name"
                                                'Cells3(RowPtr, 2) = CollectionLastChangeDate & "&nbsp;"
                                                'Cells3(RowPtr, 3) = CollectionDescription & "&nbsp;"
                                            Else
                                                If CollectionGUID = "" Then
                                                    'Cells3(RowPtr, 0) = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""0"" disabled>"
                                                    'Cells3(RowPtr, 1) = CollectionName & " (no guid)"
                                                    'Cells3(RowPtr, 2) = CollectionLastChangeDate & "&nbsp;"
                                                    'Cells3(RowPtr, 3) = CollectionDescription & "&nbsp;"
                                                Else
                                                    Dim cs As CPCSBaseClass = cp.CSNew
                                                    Dim modifiedDate As Date = Date.MinValue
                                                    Dim IsOnSite As Boolean = False
                                                    If cs.Open("Add-on Collections", GuidFieldName & "=" & cp.Db.EncodeSQLText(CollectionGUID), "", True, "ID,ModifiedDate") Then
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
                                                    ElseIf (modifiedDate >= CollectionLastChangeDate) Then
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
                                                Cell = Replace(Cell, "##date##", CollectionLastChangeDateCaption)
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

                            'BodyHTML = AdminUI.GetEditPanel(main, True, "Add-on Collection Library", "Select Add-ons to install from the Contensive Add-on Library. Please select only one at a time. Click OK to install the selected Add-on. The site may need to be stopped during the installation, but will be available again in approximately one minute.", BodyHTML)
                            'BodyHTML = BodyHTML & cp.Html.Hidden("AOCnt", RowPtr)
                            'Call main.AddLiveTabEntry("<NOBR>Collection&nbsp;Library</NOBR>", BodyHTML, "ccAdminTab")
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
                    '
                    form.title = "Add-on Library"
                    form.description = ""
                    If Not DbUpToDate Then
                        form.description = form.description & "<div style=""Margin-left:50px"">The Add-on Manager is disabled because this site's Database needs to be upgraded.</div>"
                    End If
                    If status <> "" Then
                        form.description = form.description & "<div style=""Margin-left:50px"">" & status & "</div>"
                    End If
                    returnResult = form.getHtml(cp)
                    ' returnResult = AdminUI.GetBody(main, Caption, ButtonList, "", False, False, Description, "", 0, Content.Text)
                    'Call main.AddPageTitle("Add-on Manager")
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return returnResult
        End Function
    End Class
End Namespace
