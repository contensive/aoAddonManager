VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FromLibraryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Public Function Execute(Csv As Object, main As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Dim AddonInstall As Object
    Dim collectionGuid As String
    Dim IISResetRequired As Boolean
    Dim RegisterList As String
    Dim ErrorMessage As String
    '
    collectionGuid = GetAddonOptionStringValue("collectionGuid", OptionString)
    Set AddonInstall = CreateObject("ccUpgrade.AddonInstallClass")
    Execute = AddonInstall.UpgradeAllAppsFromLibCollection(collectionGuid, main.ApplicationName, IISResetRequired, RegisterList, ErrorMessage)
    '
    Exit Function
ErrorTrap:
    If main Is Nothing Then
        Call HandleError2("(main is nothing so app.name unknown)", "error trap", "ccWeb3", "AddonManClass", "", Err.Number, Err.Source, Err.Description, True, False, "unknown")
    Else
        Call HandleError2(main.ApplicationName, "error trap", "ccWeb3", "AddonManClass", "", Err.Number, Err.Source, Err.Description, True, False, main.ServerLink)
    End If
End Function
'
''
''==========================================================================================================================================
''   Addon Manager
''       This is a form that lets you upload an addon
''       Eventually, this should be substituted with a "Addon Manager Addon" - so the interface can be improved with Contensive recompile
''==========================================================================================================================================
''
'Private Function installCollection() As String
'    On Error GoTo ErrorTrap
'    '
'    Const InstallFolderName = "install"
'    '
'    Dim showAddon As Boolean
'    Dim cellTemplate As String
'    Dim Cell As String
'    Dim CollectionCheckbox As String
'    'Dim IsOnServer As Boolean
'    Dim IsOnSite As Boolean
'    'Dim LocalCollectionXML As String
'    Dim DisplaySystem As Boolean
'    Dim CollectionSystem As Boolean
'    Dim DbUpToDate As Boolean
'    Dim GuidFieldName As String
'    Dim InstalledCollectionID As Long
'    Dim InstalledCollectionGuid As String
'    Dim PageNode As IXMLDOMNode
'    Dim SettingNode As IXMLDOMNode
'    Dim DateValue As Date
'    Dim ErrorMessage As String
'    'Dim OnServerGuidList As String
'    Dim UpgradeOK As Boolean
'    Dim RegisterList As String
'    'Dim LocalCollections As DOMDocument
'    Dim LibCollections As DOMDocument
'    Dim InstallFolder As String
'    Dim LibGuids() As String
'    Dim IISResetRequired As Boolean
'    Dim AddonNavigatorID As Long
'    Dim TargetCollectionName As String
'    Dim SQL As String
'    Dim CSAddons As Long
'    Dim addonid As Long
'    Dim LoopPtr As Long
'    Dim InstallSourcePath As String
'    Dim Deletes() As DeleteType
'    Dim DeleteCnt As Long
'    Dim CollectionName As String
'    Dim collectionGuid As String
'    Dim CollectionVersion As String
'    Dim CollectionDescription As String
'    Dim CollectionImageLink As String
'    Dim CollectionContensiveVersion As String
'    Dim CollectionLastChangeDate As String
'    'Dim Cells2() As String
'    'Dim Cells3() As String
'    Dim DeletePtr As Long
'    Dim ParentID As Long
'    Dim TargetNameSpace As String
'    Dim FormInput As String
'    Dim TargetName As String
'    Dim TargetPtr As Long
'    Dim TargetCnt As Long
'    Dim MenuCnt As Long
'    Dim NavigatorCnt As Long
'    Dim AddonCnt As Long
'    Dim AddonInUse As Boolean
'    Dim CollectionID As Long
'    Dim cnt As Long
'    Dim Ptr As Long
'    Dim UploadTab As New fastStringClass
'    Dim ModifyTab As New fastStringClass
'    Dim RowPtr As Long
'    Dim RowCnt As Long
'    Dim Body As fastStringClass
'    Dim fs As New kmaFileSystem3.FileSystemClass
'    Dim GuidGenerator As GuidGenerator
'    Dim Cells() As String
'    Dim PageNumber As Long
'    Dim ColumnCnt As Long
'    Dim FormDescription As String
'    Dim ColCaption() As String
'    Dim ColAlign() As String
'    Dim ColWidth() As String
'    Dim ColSortable() As Boolean
'    Dim PreTableCopy As String
'    Dim PostTableCopy As String
'    Dim DataRowCount As Long
'    Dim ClassStyle As String
'    Dim DefaultSortColumnPtr As Long
'    Dim BodyHTML As String
'    Dim CS As Long
'    Dim Criteria As String
'    Dim IsFound As Boolean
'    Dim AOName As String
'    Dim FieldName As String
'    Dim FieldValue As String
'    Dim UserError As String
'    Dim CDefContent As String
'    Dim Temp As String
'    Dim Content As New fastStringClass
'    Dim ButtonBar As String
'    Dim Button As String
'    Dim AdminUI As Object
'    Dim Caption As String
'    Dim Description As String
'    Dim ButtonList As String
'    Dim CollectionFilename As String
'    Dim CollectionFilePathPage As String
'    Dim Uploads() As String
'    Dim UploadsCnt As Long
'    Dim UploadPathPage As String
'    Dim CDef_Node As IXMLDOMNode
'    Dim CDefChildNode As IXMLDOMNode
'    Dim DataSourcename As String
'    Dim Doc As New DOMDocument
'    Dim InterfaceNode As IXMLDOMNode
'    Dim CollectionNode As IXMLDOMNode
'    Dim CDefNode As IXMLDOMNode
'    Dim status As String
'    Dim AllowInstallFromFolder As Boolean
'    Dim InstallLibCollectionList As String
'    Dim CollectionFile As String
'    Dim TargetCollectionID As Long
'    Dim TargetCollectionPtr As Long
'    Dim CollectionPtr As Long
'    Dim SearchCCnt As Long
'    Dim SearchCPtr As Long
'    Dim SearchAPtr As Long
'    Dim SearchMPtr As Long
'    Dim UseGUID As Boolean
'    Dim TargetAddonCnt As Long
'    Dim TargetAddonPtr As Long
'    Dim TargetAddonName As String
'    Dim TargetAddonGUID As String
'    Dim KeepTarget As Boolean
'    'Dim AddonInstall As New AddonInstallClass
'    Dim InstallPath As String
'    Dim SiteKey As String
'    Dim CollectionHelpLink As String
'    Dim CollectionDemoLink As String
'    Dim hint As String
'    '
'    'Set AddonInstall = CreateObject("ccUpgrade.AddonInstallClass")
'    'Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
'    '
'    ' workaround for migration to 4.2 - calls to cp and cpCom
'    '
'Call HandleClassAppendLogfile("AddonManager 100", "debugging")
'
'    Dim cmc As Object
'    Dim useCmc As Boolean
'    useCmc = False
'    If main.ContentServerVersion > "4.2.100" Then
'        useCmc = True
'        Set cmc = main.getCmc()
'        ' call admin UI proxy
'        Set AdminUI = CreateObject("ccWeb42.AdminUIClass")
'    Else
'        Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
'    End If
'    '
'    ' Sitekey - this is the site's identifyer back at the support site. It initially is being used
'    ' to identify site support agreements for add-on use.
'    '
'Call HandleClassAppendLogfile("AddonManager 200", "debugging")
'    SiteKey = main.getsiteproperty("sitekey", "")
'    If SiteKey = "" Then
'        Set GuidGenerator = New GuidGenerator
'        SiteKey = GuidGenerator.CreateGUID("")
'        Call main.setsiteproperty("sitekey", SiteKey)
'        Set GuidGenerator = Nothing
'    End If
'    '
'Call HandleClassAppendLogfile("AddonManager 300", "debugging")
'    DbUpToDate = True
'
'    AllowInstallFromFolder = False
'    If main.SiteProperty_BuildVersion > "3.4.060" Then
'        GuidFieldName = "ccguid"
'    Else
'        GuidFieldName = "aoguid"
'    End If
'    If Button = ButtonCancel Then
'        '
'        ' ----- redirect back to the root
'        '
'        Call main.Redirect(main.SiteProperty_AdminURL)
'    Else
'        If Not main.IsAdmin Then
'            '
'            ' ----- Put up error message
'            '
'            ButtonList = ButtonCancel
'            Content.Add AdminUI.GetFormBodyAdminOnly(main)
'        Else
'            '
'            PreTableCopy = "Use this form to upload an add-on collection. If the GUID of the add-on matches one already installed on this server, it will be updated. If the GUID is new, it will be added."
'            InstallFolder = "CollectionUpload" & CStr(GetRandomInteger())
'            InstallPath = main.PhysicalFilePath & InstallFolder & "\"
'            If (Button = ButtonOK) Then
'                '
'                '---------------------------------------------------------------------------------------------
'                ' Download and install Collections from the Collection Library
'                '---------------------------------------------------------------------------------------------
'                '
'                If main.GetStreamText("LibraryRow") <> "" Then
'                    Ptr = main.GetStreamInteger("LibraryRow")
'                    collectionGuid = main.GetStreamText("LibraryRowguid" & Ptr)
'                    InstallLibCollectionList = InstallLibCollectionList & "," & collectionGuid
'                End If
'                '
'                '---------------------------------------------------------------------------------------------
'                ' Delete collections
'                '   Before deleting each addon, make sure it is not in another collection
'                '---------------------------------------------------------------------------------------------
'                '
'                cnt = main.GetStreamInteger("accnt")
'Call HandleClassAppendLogfile("AddonManager 400", "debugging")
'                If cnt > 0 Then
'                    For Ptr = 0 To cnt - 1
'                        If main.GetStreamBoolean2("ac" & Ptr) Then
'                            TargetCollectionID = main.GetStreamInteger("acID" & Ptr)
'                            TargetCollectionName = main.GetRecordName("Add-on Collections", TargetCollectionID)
'                            '
'                            ' Clean up rules associating this collection to other objects
'                            '
'                            Call main.DeleteContentRecords("Add-on Collection CDef Rules", "collectionid=" & TargetCollectionID)
'                            Call main.DeleteContentRecords("Add-on Collection Module Rules", "collectionid=" & TargetCollectionID)
'                            '
'                            ' Delete any addons from this collection
'                            '
'                            CS = main.Opencscontent("add-ons", "collectionid=" & TargetCollectionID)
'                            Do While main.iscsok(CS)
'                                '
'                                ' Clean up the rules that might have pointed to the addon
'                                '
'                                addonid = main.getcsinteger(CS, "id")
'                                Call main.DeleteContentRecords("Admin Menuing", "addonid=" & addonid)
'                                Call main.DeleteContentRecords("Shared Styles Add-on Rules", "addonid=" & addonid)
'                                Call main.DeleteContentRecords("Add-on Scripting Module Rules", "addonid=" & addonid)
'                                Call main.DeleteContentRecords("Add-on Include Rules", "addonid=" & addonid)
'                                Call main.DeleteContentRecords("Add-on Include Rules", "includedaddonid=" & addonid)
'                                Call main.nextcsrecord(CS)
'                            Loop
'                            Call main.DeleteContentRecords("add-ons", "collectionid=" & TargetCollectionID)
'                            '
'                            ' Delete the navigator entry for the collection under 'Add-ons'
'                            '
'                            If TargetCollectionID > 0 Then
'                                AddonNavigatorID = 0
'                                CS = main.Opencscontent("Navigator Entries", "name='Manage Add-ons' and ((parentid=0)or(parentid is null))")
'                                If main.iscsok(CS) Then
'                                    AddonNavigatorID = main.getcsinteger(CS, "ID")
'                                End If
'                                Call main.closecs(CS)
'                                If AddonNavigatorID > 0 Then
'                                    Call installCollection_DeleteNavigatorBranch(TargetCollectionName, AddonNavigatorID)
'                                End If
'                                '
'                                ' Now delete the Collection record
'                                '
'                                Call main.DeleteContentRecord("Add-on Collections", TargetCollectionID)
'                                '
'                                ' Delete Navigator Entries set as installed by the collection (this may be all that is needed)
'                                '
'                                Call main.DeleteContentRecords("Navigator Entries", "installedbycollectionid=" & TargetCollectionID)
'                            End If
'                        End If
'                    Next
'                End If
'                '
'                '---------------------------------------------------------------------------------------------
'                ' Delete Add-ons
'                '---------------------------------------------------------------------------------------------
'                '
'                cnt = main.GetStreamInteger("aocnt")
'                If cnt > 0 Then
'                    For Ptr = 0 To cnt - 1
'                        If main.GetStreamBoolean2("ao" & Ptr) Then
'                            Call main.DeleteContentRecord("Add-ons", main.GetStreamInteger("aoID" & Ptr))
'                        End If
'                    Next
'                End If
'                '
'                '---------------------------------------------------------------------------------------------
'                ' Reinstall core collection
'                '---------------------------------------------------------------------------------------------
'                '
'                If main.IsDeveloper() And main.GetStreamBoolean("InstallCore") Then
'                    Call main.DeleteContentRecords("Add-on Collections", "ccguid='{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}'")
'                    If useCmc Then
'                        UpgradeOK = cmc.UpgradeAllAppsFromLibCollection("{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}", IISResetRequired, RegisterList, ErrorMessage)
'                    Else
'                        UpgradeOK = AddonInstall.UpgradeAllAppsFromLibCollection("{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}", main.ApplicationName, IISResetRequired, RegisterList, ErrorMessage)
'                    End If
'                End If
'                '
'                '---------------------------------------------------------------------------------------------
'                ' Upload new collection files
'                '---------------------------------------------------------------------------------------------
'                '
'                CollectionFilePathPage = main.ProcessFormInputFile("MetaFile", InstallFolder)
'                '
'                ' Process the MetaFile
'                '
'                If CollectionFilePathPage <> "" Then
'                    CollectionFilename = Mid(Replace(CollectionFilePathPage, InstallFolder, ""), 2)
'                    status = status & "<BR>Uploaded collection file [" & CollectionFilename & "]"
'                    '
'                    Call AppendLogFile2(main.ApplicationName, "Uploading new collection file, member=" & main.MemberName & " (" & main.memberID & "), CollectionFilename=" & CollectionFilename, App.EXEName, "AdminClass", "installCollection", 0, "", "", False, True, main.ServerLink, "AddonInstall", "")
'                    '
'                    UploadsCnt = main.GetStreamInteger("UploadCount")
'                    ReDim Uploads(UploadsCnt)
'                    For Ptr = 0 To UploadsCnt - 1
'                        UploadPathPage = main.ProcessFormInputFile("Upload" & Ptr, InstallFolder)
'                        If UploadPathPage <> "" Then
'                            Uploads(Ptr) = Mid(Replace(UploadPathPage, InstallFolder, ""), 2)
'                            Call HandleClassAppendLogfile("AddonManager", " app=" & main.ApplicationName & ", Member=" & main.MemberName & " (" & main.memberID & "), Uploads=" & Uploads(Ptr))
'                            status = status & "<BR>Uploaded collection file [" & Uploads(Ptr) & "]"
'                        End If
'                    Next
'                    AllowInstallFromFolder = True
'                    status = status & "<BR>Submitted Collection for import."
'                End If
'            End If
'            '
'            ' --------------------------------------------------------------------------------
'            '   Install Library Collections
'            ' --------------------------------------------------------------------------------
'            '
'Call HandleClassAppendLogfile("AddonManager 500", "debugging")
'            If InstallLibCollectionList <> "" Then
'                InstallLibCollectionList = Mid(InstallLibCollectionList, 2)
'                LibGuids = Split(InstallLibCollectionList, ",")
'                cnt = UBound(LibGuids) + 1
'                For Ptr = 0 To cnt - 1
'                    RegisterList = ""
'                    If Not UpgradeOK Then
'                        '
'                        ' block the reset because we will loose the error message
'                        '
'                        IISResetRequired = False
'                        main.AddUserError ("This Add-on Collection did not install correctly, " & ErrorMessage)
'                    Else
'                        '
'                        ' Save the first collection as the installed collection
'                        '
'                        If InstalledCollectionGuid = "" Then
'                            InstalledCollectionGuid = LibGuids(Ptr)
'                        End If
'                    End If
'                Next
'            End If
'            '
'            ' --------------------------------------------------------------------------------
'            '   Install Manual Collections
'            ' --------------------------------------------------------------------------------
'            '
'            If AllowInstallFromFolder Then
'                'InstallFolder = Main.PhysicalFilePath & InstallFolderName & "\"
'                If fs.CheckFileFolder(InstallPath) Then
'                    UpgradeOK = AddonInstall.InstallCollectionFilesFromFolder3(InstallPath, IISResetRequired, main.ApplicationName, ErrorMessage, InstalledCollectionGuid)
'                    If Not UpgradeOK Then
'                        If ErrorMessage = "" Then
'                            main.AddUserError ("The Add-on Collection did not install correctly, but no detailed error message was given.")
'                        Else
'                            main.AddUserError ("The Add-on Collection did not install correctly, " & ErrorMessage)
'                        End If
'                    End If
'                End If
'            End If
'            '
'            ' and delete the install folder if it was created
'            '
'            If fs.CheckFileFolder(InstallPath) Then
'                Call fs.DeleteFileFolder(InstallPath)
'            End If
'            '
'            ' --------------------------------------------------------------------------------
'            ' Get the InstalledCollectionID from the InstalledCollectionGuid
'            ' --------------------------------------------------------------------------------
'            '
'            If InstalledCollectionGuid <> "" Then
'                CS = main.Opencscontent("Add-on Collections", GuidFieldName & "=" & KmaEncodeSQLText(InstalledCollectionGuid))
'                If main.iscsok(CS) Then
'                    InstalledCollectionID = main.getcsinteger(CS, "ID")
'                End If
'                Call main.closecs(CS)
'            End If
'            '
'            ' --------------------------------------------------------------------------------
'            '   Register ActiveX files
'            ' --------------------------------------------------------------------------------
'            '
'            If RegisterList <> "" Then
'                Call AddonInstall.RegisterActiveXFiles(RegisterList)
'                Call AddonInstall.RegisterDotNet(RegisterList)
'                RegisterList = ""
'            End If
'            '
'            ' --------------------------------------------------------------------------------
'            '   IISReset if required
'            ' --------------------------------------------------------------------------------
'            '
'            If IISResetRequired And (Not main.IsUserError) Then
'                '
'                ' registers are async - make sure they are done before reset
'                '
'                Call main.IISReset
'                installCollection = "<div style=""top-margin:50px;text-align:center"">Your system will be reset to complete the installation.</div>"
'                Call main.setvisitproperty("RunOnce HelpCollectionID", InstalledCollectionID)
'                Exit Function
'            End If
'            '
'            ' --------------------------------------------------------------------------------
'            '   Forward to help page
'            ' --------------------------------------------------------------------------------
'            '
'            If (InstalledCollectionID <> 0) And (Not main.IsUserError) Then
'                Call main.setvisitproperty("RunOnce HelpCollectionID", InstalledCollectionID)
'                Call main.Redirect(main.SiteProperty_AdminURL)
'            End If
'            '
'            ' --------------------------------------------------------------------------------
'            ' Get Form
'            ' --------------------------------------------------------------------------------
'            ' Get the Collection Library tab
'            ' --------------------------------------------------------------------------------
'            '
'            ColumnCnt = 4
'            PageNumber = 1
'            ReDim ColCaption(3)
'            ReDim ColAlign(3)
'            ReDim ColWidth(3)
'            ReDim ColSortable(3)
'            'ReDim Cells3(10, 4)
'            '
'            ColCaption(0) = "Install"
'            ColAlign(0) = "center"
'            ColWidth(0) = "50"
'            ColSortable(0) = False
'            '
'            ColCaption(1) = "Name"
'            ColAlign(1) = "left"
'            ColWidth(1) = "200"
'            ColSortable(1) = False
'            '
'            ColCaption(2) = "Last&nbsp;Updated"
'            ColAlign(2) = "right"
'            ColWidth(2) = "200"
'            ColSortable(2) = False
'            '
'            ColCaption(3) = "Description"
'            ColAlign(3) = "left"
'            ColWidth(3) = "99%"
'            ColSortable(3) = False
'            '
'            'Set LocalCollections = New DOMDocument
'            'LocalCollectionXML = AddonInstall.GetConfig()
'            'Call LocalCollections.loadXML(LocalCollectionXML)
''            For Each CDef_Node In LocalCollections.documentElement.childNodes
''                If LCase(CDef_Node.baseName) = "collection" Then
''                    For Each CollectionNode In CDef_Node.childNodes
''                        If LCase(CollectionNode.baseName) = "guid" Then
''                            OnServerGuidList = OnServerGuidList & "," & CollectionNode.Text
''                            Exit For
''                        End If
''                    Next
''                End If
''            Next
'            '
'            Set LibCollections = New DOMDocument
'            'LibCollections.Load ("http://jay2.kma.net/GetCollectionList?iv=" & Main.ContentServerVersion)
'            Call LibCollections.Load("http://support.contensive.com/GetCollectionList?iv=" & main.ContentServerVersion & "&key=" & kmaEncodeURL(SiteKey) & "&name=" & kmaEncodeURL(main.ApplicationName) & "&primaryDomain=" & kmaEncodeURL(main.ServerDomainPrimary))
'            Ptr = 0
'            Do While LibCollections.readyState <> 4 And Ptr < 100
'                Sleep (100)
'                DoEvents
'                Ptr = Ptr + 1
'            Loop
'Call HandleClassAppendLogfile("AddonManager 600", "debugging")
'            If LibCollections.readyState <> 4 Then
'                UserError = "There was an error reading the Collection Library. The library website took too long to respond. Please try again later."
'                Call HandleClassAppendLogfile("AddonManager", UserError)
'                status = status & "<BR>" & UserError
'                main.AddUserError (UserError)
'            ElseIf LibCollections.parseError.errorCode <> 0 Then
'                UserError = "There was an error reading the Collection Library. The site may be unavailable. Please try again. The error was [" & LibCollections.parseError.reason & ", line " & LibCollections.parseError.Line & ", character " & LibCollections.parseError.linepos & "]"
'                Call HandleClassAppendLogfile("AddonManager", UserError)
'                status = status & "<BR>" & UserError
'                main.AddUserError (UserError)
'            Else
'               If LCase(LibCollections.documentElement.baseName) <> LCase(CollectionListRootNode) Then
'                    UserError = "There was an error reading the Collection Library file. The '" & CollectionListRootNode & "' element was not found."
'                    Call HandleClassAppendLogfile("AddonManager", UserError)
'                    status = status & "<BR>" & UserError
'                    main.AddUserError (UserError)
'                Else
'                    '
'                    ' Go through file to validate the XML, and build status message -- since service process can not communicate to user
'                    '
'                    RowPtr = 0
'                    'Content = ""
'                    cellTemplate = getLayout(guidAddonManagerLibraryListCell, "Addon Manager Library List Cell")
'                    For Each CDef_Node In LibCollections.documentElement.childNodes
'                        Cell = cellTemplate
'                        CollectionImageLink = ""
'                        CollectionCheckbox = ""
'                        CollectionName = ""
'                        CollectionLastChangeDate = ""
'                        CollectionDescription = ""
'                        CollectionHelpLink = ""
'                        CollectionDemoLink = ""
'                        Select Case LCase(CDef_Node.baseName)
'                            Case "collection"
'                                '
'                                ' Read the collection
'                                '
'                                For Each CollectionNode In CDef_Node.childNodes
'                                    Select Case LCase(CollectionNode.baseName)
'                                        Case "name"
'                                            '
'                                            ' Name
'                                            '
'                                            CollectionName = CollectionNode.Text
'                                        Case "helplink"
'                                            '
'                                            ' helpLink
'                                            '
'                                            CollectionHelpLink = CollectionNode.Text
'                                        Case "demolink"
'                                            '
'                                            ' demoLink
'                                            '
'                                            CollectionDemoLink = CollectionNode.Text
'                                        Case "guid"
'                                            '
'                                            ' Guid
'                                            '
'                                            collectionGuid = CollectionNode.Text
'                                        Case "version"
'                                            '
'                                            ' Version
'                                            '
'                                            CollectionVersion = CollectionNode.Text
'                                        Case "description"
'                                            '
'                                            ' Version
'                                            '
'                                            CollectionDescription = CollectionNode.Text
'                                        Case "imagelink"
'                                            '
'                                            ' Version
'                                            '
'                                            CollectionImageLink = CollectionNode.Text
'                                        Case "contensiveversion"
'                                            '
'                                            ' Version
'                                            '
'                                            CollectionContensiveVersion = CollectionNode.Text
'                                        Case "lastchangedate"
'                                            '
'                                            ' Version
'                                            '
'                                            CollectionLastChangeDate = CollectionNode.Text
'                                            If IsDate(CollectionLastChangeDate) Then
'                                                DateValue = CDate(CollectionLastChangeDate)
'                                                CollectionLastChangeDate = CStr(Int(DateValue))
'                                            End If
'                                            If CollectionLastChangeDate = "" Then
'                                                CollectionLastChangeDate = "unknown"
'                                            End If
'                                    End Select
'                                Next
'                                If CollectionImageLink = "" Then
'                                    CollectionImageLink = "/addonManager/libraryNoImage.jpg"
'                                End If
'                                If CollectionLastChangeDate = "" Then
'                                    CollectionLastChangeDate = "unknown"
'                                End If
'                                If CollectionDescription = "" Then
'                                    CollectionDescription = "No description is available for this add-on collection."
'                                End If
'                                'If RowPtr >= UBound(Cells3, 1) Then
'                                '    ReDim Preserve Cells3(RowPtr + 100, ColumnCnt)
'                                'End If
'                                showAddon = False
'                                If CollectionName = "" Then
'                                    'Cells3(RowPtr, 0) = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""0"" disabled>"
'                                    'Cells3(RowPtr, 1) = "no name"
'                                    'Cells3(RowPtr, 2) = CollectionLastChangeDate & "&nbsp;"
'                                    'Cells3(RowPtr, 3) = CollectionDescription & "&nbsp;"
'                                Else
'                                    If collectionGuid = "" Then
'                                        'Cells3(RowPtr, 0) = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""0"" disabled>"
'                                        'Cells3(RowPtr, 1) = CollectionName & " (no guid)"
'                                        'Cells3(RowPtr, 2) = CollectionLastChangeDate & "&nbsp;"
'                                        'Cells3(RowPtr, 3) = CollectionDescription & "&nbsp;"
'                                    Else
'                                        'IsOnServer = kmaEncodeBoolean(InStr(1, OnServerGuidList, CollectionGUID, vbTextCompare))
'                                        CS = main.Opencscontent("Add-on Collections", GuidFieldName & "=" & KmaEncodeSQLText(collectionGuid), , , , , "ID")
'                                        IsOnSite = main.iscsok(CS)
'                                        Call main.closecs(CS)
'                                        If IsOnSite Then
'                                            '
'                                            ' Already installed
'                                            '
'                                            showAddon = True
'                                            CollectionCheckbox = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""1"" disabled>&nbsp;Already installed."
'                                            'Cells3(RowPtr, 0) = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""1"" disabled>"
'                                            'Cells3(RowPtr, 1) = CollectionName & "&nbsp;(installed already)"
'                                            'Cells3(RowPtr, 2) = CollectionLastChangeDate & "&nbsp;"
'                                            'Cells3(RowPtr, 3) = CollectionDescription & "&nbsp;"
'                                        ElseIf ((CollectionContensiveVersion <> "") And (CollectionContensiveVersion > main.ContentServerVersion)) Then
'                                            '
'                                            ' wrong version
'                                            '
'                                            showAddon = True
'                                            CollectionCheckbox = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""0"" disabled>&nbsp;Disabled because this server needs to be upgraded. Contensive v" & CollectionContensiveVersion & " is required."
'                                            'Cells3(RowPtr, 0) = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""0"" disabled>"
'                                            'Cells3(RowPtr, 1) = CollectionName & "&nbsp;(Contensive v" & CollectionContensiveVersion & " needed)"
'                                            'Cells3(RowPtr, 2) = CollectionLastChangeDate & "&nbsp;"
'                                            'Cells3(RowPtr, 3) = CollectionDescription & "&nbsp;"
'                                        ElseIf Not DbUpToDate Then
'                                            '
'                                            ' Site needs to by upgraded
'                                            '
'                                            showAddon = True
'                                            CollectionCheckbox = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""0"" disabled>&nbsp;Disabled because this website database needs to be upgraded."
'                                            CollectionName = CollectionName
'                                            'Cells3(RowPtr, 0) = "<input TYPE=""CheckBox"" NAME=""LibraryRow" & RowPtr & """ VALUE=""0"" disabled>"
'                                            'Cells3(RowPtr, 1) = CollectionName & "&nbsp;(install disabled)"
'                                            'Cells3(RowPtr, 2) = CollectionLastChangeDate & "&nbsp;"
'                                            'Cells3(RowPtr, 3) = CollectionDescription & "&nbsp;"
'                                        Else
'                                            '
'                                            ' Not installed yet
'                                            '
'                                            showAddon = True
'                                            CollectionCheckbox = "<input TYPE=""CheckBox"" NAME=""LibraryRow"" VALUE=""" & RowPtr & """ onClick=""clearLibraryRows('" & RowPtr & "');"">" & main.GetFormInputHidden("LibraryRowGuid" & RowPtr, collectionGuid) & main.GetFormInputHidden("LibraryRowName" & RowPtr, CollectionName) & "&nbsp;Install"
'                                            'Cells3(RowPtr, 0) = "<input TYPE=""CheckBox"" NAME=""LibraryRow"" VALUE=""" & RowPtr & """ onClick=""clearLibraryRows('" & RowPtr & "');"">" & Main.GetFormInputHidden("LibraryRowGuid" & RowPtr, CollectionGUID) & Main.GetFormInputHidden("LibraryRowName" & RowPtr, CollectionName)
'                                            'Cells3(RowPtr, 1) = CollectionName & "&nbsp;"
'                                            'Cells3(RowPtr, 2) = CollectionLastChangeDate & "&nbsp;"
'                                            'Cells3(RowPtr, 3) = CollectionDescription & "&nbsp;"
'                                        End If
'                                    End If
'                                End If
'                                If CollectionDemoLink <> "" Then
'                                    CollectionDescription = CollectionDescription & "<div class=""amDemoLink""><a target=""_blank"" href=""" & CollectionDemoLink & """>Demo</a></div>"
'                                End If
'                                If CollectionHelpLink <> "" Then
'                                    CollectionDescription = CollectionDescription & "<div class=""amHelpLink""><a target=""_blank"" href=""" & CollectionHelpLink & """>Reference</a></div>"
'                                End If
'                                If showAddon Then
'                                    Cell = Replace(Cell, "##imageLink##", CollectionImageLink)
'                                    Cell = Replace(Cell, "##checkbox##", CollectionCheckbox)
'                                    Cell = Replace(Cell, "##name##", CollectionName)
'                                    Cell = Replace(Cell, "##date##", CollectionLastChangeDate)
'                                    Cell = Replace(Cell, "##description##", CollectionDescription)
'                                    BodyHTML = BodyHTML & Cell
'                                End If
'                                RowPtr = RowPtr + 1
'                        End Select
'                    Next
'                End If
'                BodyHTML = "" _
'                    & cr & "<script language=""JavaScript"">" _
'                    & "function clearLibraryRows(r) {" _
'                    & "var c,p;" _
'                    & "c=document.getElementsByName('LibraryRow');" _
'                        & "for (p=0;p<c.length;p++){" _
'                            & "if(c[p].value!=r)c[p].checked=false;" _
'                        & "}" _
'                    & "" _
'                    & "}" _
'                    & "</script>" _
'                    & "<input type=hidden name=LibraryCnt value=""" & RowPtr & """>" _
'                    & cr & "<div style=""width:100%"">" _
'                    & kmaIndent(BodyHTML) _
'                    & cr & "</div>" _
'                    & ""
'                'BodyHTML = "" _
'                '    & "<script language=""JavaScript"">" _
'                '    & "function clearLibraryRows(r) {" _
'                '    & "var c,p;" _
'                '    & "c=document.getElementsByName('LibraryRow');" _
'                '        & "for (p=0;p<c.length;p++){" _
'                '            & "if(c[p].value!=r)c[p].checked=false;" _
'                '        & "}" _
'                '    & "" _
'                '    & "}" _
'                '    & "</script>" _
'                '    & "<input type=hidden name=LibraryCnt value=""" & RowPtr & """>" _
'                '    & "<div style=""width:100%"">" & AdminUI.GetReport2(Main, RowPtr, ColCaption, ColAlign, ColWidth, Cells3, RowPtr, 1, "", PostTableCopy, RowPtr, "ccAdmin", ColSortable, 0) & "</div>" _
'                '    & ""
'                BodyHTML = AdminUI.GetEditPanel(main, True, "Add-on Collection Library", "Select Add-ons to install from the Contensive Add-on Library. Please select only one at a time. Click OK to install the selected Add-on. The site may need to be stopped during the installation, but will be available again in approximately one minute.", BodyHTML)
'                BodyHTML = BodyHTML & main.GetFormInputHidden("AOCnt", RowPtr)
'                Call main.AddLiveTabEntry("<NOBR>Collection&nbsp;Library</NOBR>", BodyHTML, "ccAdminTab")
'            End If
'Call HandleClassAppendLogfile("AddonManager 700", "debugging")
'            '
'            ' --------------------------------------------------------------------------------
'            ' Current Collections Tab
'            ' --------------------------------------------------------------------------------
'            '
'            ColumnCnt = 2
'            PageNumber = 1
'            ReDim ColCaption(2)
'            ReDim ColAlign(2)
'            ReDim ColWidth(2)
'            ReDim ColSortable(2)
'            '
'            ColCaption(0) = "Del"
'            ColAlign(0) = "center"
'            ColWidth(0) = "50px"
'            ColSortable(0) = False
'            '
'            ColCaption(1) = "Name"
'            ColAlign(1) = "left"
'            ColWidth(1) = ""
'            ColSortable(1) = False
'            '
'            DisplaySystem = False
'            If main.SiteProperty_BuildVersion < "3.4.139" Then
'                '
'                ' before system attribute
'                '
'                CS = main.Opencscontent("Add-on Collections", , "Name")
'            ElseIf Not main.IsDeveloper Then
'                '
'                ' non-developers
'                '
'                CS = main.Opencscontent("Add-on Collections", "((system is null)or(system=0))", "Name")
'            Else
'                '
'                ' developers
'                '
'                DisplaySystem = True
'                CS = main.Opencscontent("Add-on Collections", , "Name")
'            End If
'            ReDim Preserve Cells(main.GetCSRowCount(CS), ColumnCnt)
'            RowPtr = 0
'            Do While main.iscsok(CS)
'                Cells(RowPtr, 0) = main.GetFormInputCheckBox("AC" & RowPtr) & main.GetFormInputHidden("ACID" & RowPtr, main.getcsinteger(CS, "ID"))
'                'Cells(RowPtr, 1) = "<a href=""" & Main.SiteProperty_AdminURL & "?id=" & Main.GetCSInteger(CS, "ID") & "&cid=" & Main.GetCSInteger(CS, "ContentControlID") & "&af=4""><img src=""/cclib/images/IconContentEdit.gif"" border=0></a>"
'                Cells(RowPtr, 1) = main.getcsText(CS, "name")
'                If DisplaySystem Then
'                    If main.GetCSBoolean(CS, "system") Then
'                        Cells(RowPtr, 1) = Cells(RowPtr, 1) & " (system)"
'                    End If
'                End If
'                Call main.nextcsrecord(CS)
'                RowPtr = RowPtr + 1
'            Loop
'            Call main.closecs(CS)
'            BodyHTML = "<div style=""width:100%"">" & AdminUI.GetReport2(main, RowPtr, ColCaption, ColAlign, ColWidth, Cells, RowPtr, 1, "", PostTableCopy, RowPtr, "ccAdmin", ColSortable, 0) & "</div>"
'            BodyHTML = AdminUI.GetEditPanel(main, True, "Add-on Collections", "Use this form to uninstall (remove) add-on collections from your site.", BodyHTML)
'            BodyHTML = BodyHTML & main.GetFormInputHidden("accnt", RowPtr)
'            BodyHTML = Replace(BodyHTML, "vertical-align:bottom;text-align:right;", "width:50px;vertical-align:bottom;text-align:right;", , , vbTextCompare)
'            Call main.AddLiveTabEntry("Uninstall&nbsp;Collections", BodyHTML, "ccAdminTab")
'            '
'            ' --------------------------------------------------------------------------------
'            ' Get the Upload Add-ons tab
'            ' --------------------------------------------------------------------------------
'            '
'            Set Body = New fastStringClass
'            If Not DbUpToDate Then
'                Call Body.Add("<p>Add-on upload is disabled because your site database needs to be updated.</p>")
'            Else
'                Call Body.Add(AdminUI.EditTableOpen)
'
'                If main.IsDeveloper Then
'                    Call Body.Add(AdminUI.GetEditRow(main, main.GetFormInputCheckBox("InstallCore"), "Reinstall Core Collection", "", False, False, ""))
'                End If
'                Call Body.Add(AdminUI.GetEditRow(main, main.GetFormInputFile("MetaFile"), "Add-on Collection File(s)", "", True, False, ""))
'                FormInput = "" _
'                    & "<TABLE id=""UploadInsert"" border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%"">" _
'                    & "</Table>" _
'                    & "<TABLE border=""0"" cellpadding=""0"" cellspacing=""1"" width=""100%"">" _
'                    & "<TR><TD align=""left""><a href=""#"" onClick=""InsertUpload(); return false;"">+ Add more files</a></TD></TR>" _
'                    & "</Table>" _
'                    & main.GetFormInputHidden("UploadCount", 1, "UploadCount") _
'                    & ""
'                Call Body.Add(AdminUI.GetEditRow(main, FormInput, "&nbsp;", "", True, False, ""))
'                Call Body.Add(AdminUI.EditTableClose)
'            End If
'            Call main.AddLiveTabEntry("Add&nbsp;Manually", AdminUI.GetEditPanel(main, True, "Install or Update an Add-on Collection.", "Use this form to upload a new or updated Add-on Collection to your site. A collection file can be a single xml configuration file, a single zip file containing the configuration file and other resource files, or a configuration with other resource files uploaded separately. Use the 'Add more files' link to add as many files as you need. When you hit OK, the Collection will be checked, and only submitted if all files are uploaded.", Body.Text), "ccAdminTab")
'            '
'            ' --------------------------------------------------------------------------------
'            ' Build Page from tabs
'            ' --------------------------------------------------------------------------------
'            '
'            Content.Add main.GetLiveTabs()
'            Content.Add main.GetFormInputHidden(RequestNameRefreshBlock, main.GetFormSN)
'            '
'            ButtonList = ButtonCancel & "," & ButtonOK
''!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'' I dont understand why this is here, unless it is just leftover from when it was integrated into the admin site
'            'Content.Add (Main.GetFormInputHidden(RequestNameAdminSourceForm, AdminFormUploadAddonPackage))
'        End If
'        '
'        ' Output the Add-on
'        '
'        Caption = "Add-on Manager"
'        Description = "<div>Use the add-on manager to add and remove Add-ons from your Contensive installation.</div>"
'        If Not DbUpToDate Then
'            Description = Description & "<div style=""Margin-left:50px"">The Add-on Manager is disabled because this site's Database needs to be upgraded.</div>"
'        End If
'        If status <> "" Then
'            Description = Description & "<div style=""Margin-left:50px"">" & status & "</div>"
'        End If
'        installCollection = AdminUI.GetBody(main, Caption, ButtonList, "", False, False, Description, "", 0, Content.Text)
'        Call main.AddPageTitle("Add-on Manager")
'    End If
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleClassTrapError("installCollection")
'End Function
''
''
''
'Private Sub installCollection_DeleteNavigatorBranch(EntryName As String, EntryParentID As Long)
'    On Error GoTo ErrorTrap
'    '
'    Dim CS As Long
'    Dim EntryID As Long
'    '
'    Call AppendLogFile2(main.ApplicationName, "Deleting Navigator Branch", App.EXEName, "AdminClass", "installCollection_DeleteNavigatorBranch", 0, "", "", False, True, main.ServerLink, "AddonInstall", "")
'    '
'    If EntryParentID = 0 Then
'        CS = main.Opencscontent("Navigator Entries", "(name=" & KmaEncodeSQLText(EntryName) & ")and((parentID is null)or(parentid=0))")
'    Else
'        CS = main.Opencscontent("Navigator Entries", "(name=" & KmaEncodeSQLText(EntryName) & ")and(parentID=" & KmaEncodeSQLNumber(EntryParentID) & ")")
'    End If
'    If main.iscsok(CS) Then
'        EntryID = main.getcsinteger(CS, "ID")
'    End If
'    Call main.closecs(CS)
'    '
'    If EntryID <> 0 Then
'        CS = main.Opencscontent("Navigator Entries", "(parentID=" & KmaEncodeSQLNumber(EntryID) & ")")
'        Do While main.iscsok(CS)
'            Call installCollection_DeleteNavigatorBranch(main.getcsText(CS, "name"), EntryID)
'            Call main.nextcsrecord(CS)
'        Loop
'        Call main.closecs(CS)
'        Call main.DeleteContentRecord("Navigator Entries", EntryID)
'    End If
'    '
'    Exit Sub
'ErrorTrap:
'    Call HandleClassTrapError("installCollection_DeleteNavigatorBranch")
'End Sub
''
''========================================================================
'' ----- Get an XML nodes attribute based on its name
''========================================================================
''
'Private Function GetXMLAttribute(found As Boolean, Node As IXMLDOMNode, Name As String, DefaultIfNotFound As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim NodeAttribute As IXMLDOMAttribute
'    Dim REsultNode As IXMLDOMNode
'    Dim UcaseName As String
'    '
'    found = False
'    Set REsultNode = Node.Attributes.getNamedItem(Name)
'    If (REsultNode Is Nothing) Then
'        UcaseName = UCase(Name)
'        For Each NodeAttribute In Node.Attributes
'            If UCase(NodeAttribute.nodeName) = UcaseName Then
'                GetXMLAttribute = NodeAttribute.nodeValue
'                found = True
'                Exit For
'                End If
'            Next
'    Else
'        GetXMLAttribute = REsultNode.nodeValue
'        found = True
'    End If
'    If Not found Then
'        GetXMLAttribute = DefaultIfNotFound
'    End If
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassTrapError("GetXMLAttribute")
'    End Function
''
''
''
'Private Sub HandleClassAppendLogfile(MethodName As String, Context As String)
'    Call AppendLogFile2(main.ApplicationName, Context, "ccWeb3", "AddonManClass", MethodName, 0, "", "", False, True, main.ServerLink, "", "trace")
'
'End Sub
''
''===========================================================================
''
''===========================================================================
''
'Private Sub HandleClassTrapError(MethodName As String, Optional Context As String)
'    '
'    If main Is Nothing Then
'        Call HandleError2("unknown", Context, "ccWeb3", "AddonManClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, "unknown")
'    Else
'        Call HandleError2(main.ApplicationName, Context, "ccWeb3", "AddonManClass", MethodName, Err.Number, Err.Source, Err.Description, True, False, main.ServerLink)
'    End If
'    '
'    End Sub
''
''
''
'Private Function GetParentIDFromNameSpace(ContentName As String, NameSpace As String) As Long
'    On Error GoTo ErrorTrap
'    '
'    Dim NameSplit() As String
'    Dim ParentNameSpace As String
'    Dim ParentName As String
'    Dim ParentID As Long
'    Dim Pos As Long
'    Dim CS As Long
'    '
'    GetParentIDFromNameSpace = 0
'    If NameSpace <> "" Then
'        'ParentName = ParentNameSpace
'        Pos = InStr(1, NameSpace, ".")
'        If Pos = 0 Then
'            ParentName = NameSpace
'            ParentNameSpace = ""
'        Else
'            ParentName = Mid(NameSpace, Pos + 1)
'            ParentNameSpace = Mid(NameSpace, 1, Pos - 1)
'        End If
'        If ParentNameSpace = "" Then
'            CS = main.Opencscontent(ContentName, "(name=" & KmaEncodeSQLText(ParentName) & ")and((parentid is null)or(parentid=0))", "ID", , , , "ID")
'            If main.iscsok(CS) Then
'                GetParentIDFromNameSpace = main.getcsinteger(CS, "ID")
'            End If
'            Call main.closecs(CS)
'        Else
'            ParentID = GetParentIDFromNameSpace(ContentName, ParentNameSpace)
'            CS = main.Opencscontent(ContentName, "(name=" & KmaEncodeSQLText(ParentName) & ")and(parentid=" & ParentID & ")", "ID", , , , "ID")
'            If main.iscsok(CS) Then
'                GetParentIDFromNameSpace = main.getcsinteger(CS, "ID")
'            End If
'            Call main.closecs(CS)
'        End If
'    End If
'    '
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassTrapError("GetParentIDFromNameSpace", False)
'End Function
''
''
''
'Private Function getLayout(layoutGuid As String, DefaultName As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim CS As Long
'    '
'    CS = main.Opencscontent("layouts", "ccguid=" & KmaEncodeSQLText(layoutGuid))
'    If Not main.iscsok(CS) Then
'        Call main.closecs(CS)
'        CS = main.InsertCSContent("layouts")
'        If main.iscsok(CS) Then
'            Call main.setcs(CS, "ccguid", layoutGuid)
'            Call main.setcs(CS, "name", DefaultName)
'            Call main.setcs(CS, "layout", "<!-- layout record created " & Now & " -->")
'        End If
'    End If
'    If main.iscsok(CS) Then
'        getLayout = main.getcs(CS, "layout")
'    End If
'    Call main.closecs(CS)
'    '
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassTrapError("getLayout", False)
'End Function
'
'
