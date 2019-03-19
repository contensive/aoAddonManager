﻿
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.addonManager
    Public Class adminUIClass
        '========================================================================
        ' This is the header file controls the online forums architecture
        '
        ' This file and its contents are copyright by Kidwell McGowan Associates.
        '========================================================================
        '
        Public Enum SortingStateEnum
            NotSortable = 0
            SortableSetAZ = 1
            SortableSetza = 2
            SortableNotSet = 3
        End Enum
        '
        Private cp As CPBaseClass
        '
        '===========================================================================
        ''' <summary>
        ''' HandleError: Page List Trap Errors
        ''' </summary>
        ''' <param name="MethodName"></param>
        Private Sub HandleTrapError(ByVal MethodName As String)
            cp.Site.ErrorReport(ex)
            '
            Call handleException(cp, New Exception("unexpected exception in method [" & MethodName & "]"))
            '
        End Sub
        '
        '===========================================================================
        '
        '===========================================================================
        '
        Public Function GetTitleBar(ByVal Title As String, ByVal Description As String) As String
            On Error GoTo ErrorTrap
            '
            Dim Copy As String
            '
            GetTitleBar = "<div class=""ccAdminTitleBar"">" & Title
            'GetTitleBar = "<div class=""ccAdminTitleBar"">" & Title & "</div>"
            Copy = Description
            If InStr(1, Copy, "<p>", vbTextCompare) = 1 Then
                Copy = Mid(Copy, 4)
                If InStrRev(Copy, "</p>", , vbTextCompare) = (Len(Copy) - 4) Then
                    Copy = Mid(Copy, 1, Len(Copy) - 4)
                End If
            End If
            '
            ' Add Errors
            '
            If cp.core.main_IsUserError Then
                Copy = Copy & "<div>" & cp.core.main_GetUserError & "</div>"
            End If
            '
            If Copy <> "" Then
                GetTitleBar = GetTitleBar & "<div>&nbsp;</div><div class=""ccAdminInfoBar ccPanel3DReverse"">" & Copy & "</div>"
            End If
            GetTitleBar = GetTitleBar & "</div>"
            '
            Exit Function
            '
            ' ----- Error Trap
            '
ErrorTrap:
            Call HandleTrapError("GetTitleBar")
        End Function
        '
        '========================================================================
        ' Get the Normal Edit Button Bar String
        '
        '   used on Normal Edit and others
        '========================================================================
        '
        Friend Function GetEditButtonBar2(ByVal MenuDepth As Integer, ByVal AllowDelete As Boolean, ByVal AllowCancel As Boolean, ByVal allowSave As Boolean, ByVal AllowSpellCheck As Boolean, ByVal AllowPublish As Boolean, ByVal AllowAbort As Boolean, ByVal AllowSubmit As Boolean, ByVal AllowApprove As Boolean, ByVal AllowAdd As Boolean, ByVal AllowReloadCDef As Boolean, ByVal HasChildRecords As Boolean, ByVal IsPageContent As Boolean, ByVal AllowMarkReviewed As Boolean, ByVal AllowRefresh As Boolean, ByVal AllowCreateDuplicate As Boolean) As String
            On Error GoTo ErrorTrap
            '
            Dim JSOnClick As String
            '
            GetEditButtonBar2 = ""
            '
            If AllowCancel Then
                If MenuDepth = 1 Then
                    '
                    ' Close if this is the root depth of a popup window
                    '
                    GetEditButtonBar2 = GetEditButtonBar2 & "<input TYPE=""SUBMIT"" NAME=""BUTTON"" VALUE=""" & ButtonClose & """ OnClick=""window.close();"">"
                Else
                    '
                    ' Cancel
                    '
                    GetEditButtonBar2 = GetEditButtonBar2 & "<input TYPE=""SUBMIT"" NAME=""BUTTON"" VALUE=""" & ButtonCancel & """ onClick=""return processSubmit(this);"">"
                End If
            End If
            If allowSave Then
                '
                ' Save
                '
                GetEditButtonBar2 = GetEditButtonBar2 & "<input TYPE=""SUBMIT"" NAME=""BUTTON"" VALUE=""" & ButtonSave & """ onClick=""return processSubmit(this);"">"
                '
                ' OK
                '
                GetEditButtonBar2 = GetEditButtonBar2 & "<input TYPE=""SUBMIT"" NAME=""BUTTON"" VALUE=""" & ButtonOK & """ onClick=""return processSubmit(this);"">"
                If AllowAdd Then
                    '
                    ' OK
                    '
                    GetEditButtonBar2 = GetEditButtonBar2 & "<input TYPE=""SUBMIT"" NAME=""BUTTON"" VALUE=""" & ButtonSaveAddNew & """ onClick=""return processSubmit(this);"">"
                End If
            End If
            If AllowDelete Then
                '
                ' Delete
                '
                If IsPageContent Then
                    JSOnClick = "if(!DeletePageCheck())return false;"
                ElseIf HasChildRecords Then
                    JSOnClick = "if(!DeleteCheckWithChildren())return false;"
                Else
                    JSOnClick = "if(!DeleteCheck())return false;"
                End If
                GetEditButtonBar2 = GetEditButtonBar2 & "<input TYPE=SUBMIT NAME=BUTTON VALUE=""" & ButtonDelete & """ onClick=""" & JSOnClick & """>"
            Else
                GetEditButtonBar2 = GetEditButtonBar2 & "<input TYPE=SUBMIT NAME=BUTTON disabled=""disabled"" VALUE=""" & ButtonDelete & """>"
            End If
            '    If AllowSpellCheck Then
            '        '
            '        ' Spell Check
            '        '
            '        GetEditButtonBar2 = GetEditButtonBar2 & "<input TYPE=""SUBMIT"" NAME=""BUTTON"" VALUE=""" & ButtonSpellCheck & """ onClick=""return processSubmit(this);"">"
            '    End If
            If AllowPublish Then
                '
                ' Publish
                '
                GetEditButtonBar2 = GetEditButtonBar2 & cp.core.main_GetFormButton(ButtonPublish, RequestNameButton)
            End If
            If AllowAbort Then
                '
                ' Abort
                '
                GetEditButtonBar2 = GetEditButtonBar2 & cp.core.main_GetFormButton(ButtonAbortEdit, RequestNameButton)
            End If
            If AllowSubmit Then
                '
                ' Submit for Publishing
                '
                GetEditButtonBar2 = GetEditButtonBar2 & cp.core.main_GetFormButton(ButtonPublishSubmit, RequestNameButton)
            End If
            If AllowApprove Then
                '
                ' Approve Publishing
                '
                GetEditButtonBar2 = GetEditButtonBar2 & cp.core.main_GetFormButton(ButtonPublishApprove, RequestNameButton)
            End If
            If AllowReloadCDef Then
                '
                ' Reload Content Definitions
                '
                GetEditButtonBar2 = GetEditButtonBar2 & cp.core.main_GetFormButton(ButtonReloadCDef, RequestNameButton)
            End If
            If AllowMarkReviewed Then
                '
                ' Reload Content Definitions
                '
                GetEditButtonBar2 = GetEditButtonBar2 & cp.core.main_GetFormButton(ButtonMarkReviewed, RequestNameButton)
            End If
            If AllowRefresh Then
                '
                ' just like a save, but don't save jsut redraw
                '
                GetEditButtonBar2 = GetEditButtonBar2 & cp.core.main_GetFormButton(ButtonRefresh, RequestNameButton)
            End If
            If AllowCreateDuplicate Then
                '
                ' just like a save, but don't save jsut redraw
                '
                GetEditButtonBar2 = GetEditButtonBar2 & cp.core.main_GetFormButton(ButtonCreateDuplicate, RequestNameButton, , "return processSubmit(this)")
            End If
            '
            GetEditButtonBar2 = "" _
                & vbCrLf & vbTab & GetHTMLComment("ButtonBar") _
                & vbCrLf & vbTab & "<div class=""ccButtonCon"">" _
                & kmaIndent(GetEditButtonBar2) _
                & vbCrLf & vbTab & "</div><!-- ButtonBar End -->" _
                & ""
            Exit Function
            '
            ' ----- Error Trap
            '
ErrorTrap:
            Call HandleTrapError("GetEditButtonBar2")
            '
        End Function
        '
        '========================================================================
        ' Return a panel header with the header message reversed out of the left
        '========================================================================
        '
        Public Function GetHeader(ByVal HeaderMessage As String, Optional ByVal RightSideMessage As String = "") As String
            On Error GoTo ErrorTrap
            '
            Dim s As String
            '
            If RightSideMessage = "" Then
                RightSideMessage = FormatDateTime(cp.core.main_PageStartTime)
            End If
            '
            If InStr(1, HeaderMessage & RightSideMessage, vbCrLf) Then
                s = "" _
                    & cr & "<td width=""50%"" valign=Middle class=""cchLeft"">" _
                    & kmaIndent(HeaderMessage) _
                    & cr & "</td>" _
                    & cr & "<td width=""50%"" valign=Middle class=""cchRight"">" _
                    & kmaIndent(RightSideMessage) _
                    & cr & "</td>"
            Else
                s = "" _
                    & cr & "<td width=""50%"" valign=Middle class=""cchLeft"">" & HeaderMessage & "</td>" _
                    & cr & "<td width=""50%"" valign=Middle class=""cchRight"">" & RightSideMessage & "</td>"
            End If
            s = "" _
                & cr & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr>" _
                & kmaIndent(s) _
                & cr & "</tr></table>" _
                & ""
            s = "" _
                & cr & "<div class=""ccHeaderCon"">" _
                & kmaIndent(s) _
                & cr & "</div>" _
                & ""
            GetHeader = s
            Exit Function
            '
ErrorTrap:
            Call HandleTrapError("GetHeader")
        End Function
        '
        '
        '
        Public Function GetButtonsFromList(ByVal ButtonList As String, ByVal AllowDelete As Boolean, ByVal AllowAdd As Boolean, ByVal ButtonName As String) As String
            On Error GoTo ErrorTrap
            '
            Dim s As String
            Dim Buttons() As String
            Dim Ptr As Integer
            '
            If Trim(ButtonList) <> "" Then
                Buttons = Split(ButtonList, ",")
                For Ptr = 0 To UBound(Buttons)
                    Select Case Trim(Buttons(Ptr))
                        Case Trim(ButtonDelete)
                            If AllowDelete Then
                                s = s & "<input TYPE=SUBMIT NAME=""" & ButtonName & """ VALUE=""" & Buttons(Ptr) & """ onClick=""if(!DeleteCheck())return false;"">"
                            Else
                                s = s & "<input TYPE=SUBMIT NAME=""" & ButtonName & """ DISABLED VALUE=""" & Buttons(Ptr) & """>"
                            End If
                        Case Trim(ButtonClose)
                            s = s & cp.core.main_GetFormButton(Buttons(Ptr), , , "window.close();")
                        Case Trim(ButtonAdd)
                            If AllowAdd Then
                                s = s & "<input TYPE=SUBMIT NAME=""" & ButtonName & """ VALUE=""" & Buttons(Ptr) & """ onClick=""return processSubmit(this);"">"
                            Else
                                s = s & "<input TYPE=SUBMIT NAME=""" & ButtonName & """ DISABLED VALUE=""" & Buttons(Ptr) & """ onClick=""return processSubmit(this);"">"
                            End If
                        Case ""
                        Case Else
                            s = s & cp.core.main_GetFormButton(Buttons(Ptr), ButtonName)
                    End Select
                Next
            End If

            GetButtonsFromList = s

            Exit Function
            '
ErrorTrap:
            Call HandleTrapError("GetButtonsFromList")
        End Function
        '
        '
        '
        Public Function GetButtonBar(ByVal LeftButtons As String, ByVal RightButtons As String) As String
            On Error GoTo ErrorTrap
            '
            If LeftButtons & RightButtons = "" Then
                '
                ' nothing there
                '
            ElseIf RightButtons = "" Then
                GetButtonBar = "" _
                    & cr & "<table border=0 cellpadding=0 cellspacing=0 width=""100%"">" _
                    & cr2 & "<tr>" _
                    & cr3 & "<td align=left  class=""ccButtonCon"">" & LeftButtons & "</td>" _
                    & cr2 & "</tr>" _
                    & cr & "</table>"
            Else
                GetButtonBar = "" _
                    & cr & "<table border=0 cellpadding=0 cellspacing=0 width=""100%"">" _
                    & cr2 & "<tr>" _
                    & cr3 & "<td align=left  class=""ccButtonCon"">" _
                    & cr4 & "<table border=0 cellpadding=0 cellspacing=0 width=""100%"">" _
                    & cr5 & "<tr>" _
                    & cr6 & "<td width=""50%"" align=left>" & LeftButtons & "</td>" _
                    & cr6 & "<td width=""50%"" align=right>" & RightButtons & "</td>" _
                    & cr5 & "</tr>" _
                    & cr4 & "</table>" _
                    & cr3 & "</td>" _
                    & cr2 & "</tr>" _
                    & cr & "</table>"
            End If
            Exit Function
            '
ErrorTrap:
            Call HandleTrapError("GetButtonBar")
        End Function
        '
        '
        '
        Friend Function GetButtonBarForIndex(ByVal LeftButtons As String, ByVal RightButtons As String, ByVal PageNumber As Integer, ByVal RecordsPerPage As Integer, ByVal PageCount As Integer) As String
            On Error GoTo ErrorTrap
            '
            Dim Ptr As Integer
            Dim Nav As String
            Dim NavStart As Integer
            Dim NavEnd As Integer
            '
            GetButtonBarForIndex = GetButtonBar(LeftButtons, RightButtons)
            NavStart = PageNumber - 9
            If NavStart < 1 Then
                NavStart = 1
            End If
            NavEnd = NavStart + 20
            If NavEnd > PageCount Then
                NavEnd = PageCount
                NavStart = NavEnd - 20
                If NavStart < 1 Then
                    NavStart = 1
                End If
            End If
            If NavStart > 1 Then
                Nav = Nav & "<li onclick=""bbj(this);"">1</li><li class=""delim"">&#171;</li>"
            End If
            For Ptr = NavStart To NavEnd
                Nav = Nav & "<li onclick=""bbj(this);"">" & Ptr & "</li>"
            Next
            If NavEnd < PageCount Then
                Nav = Nav & "<li class=""delim"">&#187;</li><li onclick=""bbj(this);"">" & PageCount & "</li>"
            End If
            Nav = Replace(Nav, ">" & PageNumber & "<", " class=""hit"">" & PageNumber & "<")
            GetButtonBarForIndex = GetButtonBarForIndex _
                & cr & "<script language=""javascript"">function bbj(p){document.getElementsByName('indexGoToPage')[0].value=p.innerHTML;document.adminForm.submit();}</script>" _
                & cr & "<div class=""ccJumpCon"">" _
                & cr2 & "<ul><li class=""caption"">Page</li>" _
                & cr3 & Nav _
                & cr2 & "</ul>" _
                & cr & "</div>"
            '    GetButtonBarForIndex = GetButtonBarForIndex _
            '        & CR & "<script language=""javascript"">function bbj(p){document.getElementsByName('indexGoToPage')[0].value=p.innerHTML;document.adminForm.submit();}</script>" _
            '        & CR & "<div class=""ccJumpCon"">" _
            '        & cr2 & "<ul>Page&nbsp;" _
            '        & cr3 & Nav _
            '        & cr2 & "</ul>" _
            '        & CR & "</div>"
            Exit Function
            '
ErrorTrap:
            Call HandleTrapError("GetButtonBar")
        End Function
        '
        '========================================================================
        '
        '========================================================================
        '
        Public Function GetBody(ByVal Caption As String, ByVal ButtonListLeft As String, ByVal ButtonListRight As String, ByVal AllowAdd As Boolean, ByVal AllowDelete As Boolean, ByVal Description As String, ByVal ContentSummary As String, ByVal ContentPadding As Integer, ByVal Content As String) As String
            On Error GoTo ErrorTrap
            '
            Dim ContentCell As String
            Dim Stream As New fastStringClass
            Dim ButtonBarLeft As String
            Dim ButtonBarRight As String
            Dim Buttons() As String
            Dim Button As String
            Dim Ptr As Integer
            Dim ButtonBar As String
            Dim ButtonList As String
            Dim Copy As String
            Dim LeftButtons As String
            Dim RightButtons As String
            Dim CellContentSummary As String
            '
            ' Build ButtonBar
            '
            If Trim(ButtonListLeft) <> "" Then
                LeftButtons = GetButtonsFromList(ButtonListLeft, AllowDelete, AllowAdd, "Button")
            End If
            If Trim(ButtonListRight) <> "" Then
                RightButtons = GetButtonsFromList(ButtonListRight, AllowDelete, AllowAdd, "Button")
            End If
            ButtonBar = GetButtonBar(LeftButtons, RightButtons)
            '    '
            '    ' ButtonBar Top
            '    '
            '    Stream.Add( ButtonBar
            '    '
            '    ' Header
            '    '
            '    Stream.Add( GetTitleBar( Caption, Description))
            '    '
            '    ' Content Summary
            '    '
            If ContentSummary <> "" Then
                CellContentSummary = "" _
                    & cr & "<div class=""ccPanelBackground"" style=""padding:10px;"">" _
                    & kmaIndent(cp.core.main_GetPanel(ContentSummary, "ccPanel", "ccPanelShadow", "ccPanelHilite", "100%", 5)) _
                    & cr & "</div>"
            End If
            '
            ' Content
            '
            '    Stream.Add( "" _
            '        & CR  & "<table border=0 cellpadding=" & ContentPadding & " cellspacing=0 width=""100%"">" _
            '        & CR  & vbTab & "<tr>" _
            '        & CR  & vbTab & vbTab & "<td style=""width:100%;vertical-align:top;text-align:left;"">" & Content & "</td>" _
            '        & CR  & vbTab & vbTab & "<td style=""width:1px;""><img src="""" width=""1"" height=""400""></td>" _
            '        & CR  & vbTab & "</tr>" _
            '        & CR  & "</table>" _
            '        & ""
            '    If ContentPadding > 0 Then
            '        Stream.Add( "<table border=0 cellpadding=" & ContentPadding & " cellspacing=0 width=""100%""><tr><td>")
            '        Stream.Add( Content)
            '        Stream.Add( "</td></tr></table>")
            '    Else
            '        Stream.Add( Content)
            '    End If
            '
            ' ButtonBar Bottom
            '
            ContentCell = "" _
                & cr & "<div style=""padding:" & ContentPadding & "px;"">" _
                & kmaIndent(Content) _
                & cr & "</div>"
            GetBody = GetBody _
                & kmaIndent(ButtonBar) _
                & kmaIndent(GetTitleBar(Caption, Description)) _
                & kmaIndent(CellContentSummary) _
                & kmaIndent(ContentCell) _
                & kmaIndent(ButtonBar) _
                & cr & cp.core.main_GetFormInputHidden(RequestNameRefreshBlock, cp.core.main_GetFormSN) _
                & ""
            '    GetBody = GetBody _
            '        & kmaIndent(ButtonBar) _
            '        & kmaIndent(GetTitleBar( Caption, Description)) _
            '        & kmaIndent(CellContentSummary) _
            '        & CR  & "<table border=0 cellpadding=" & ContentPadding & " cellspacing=0 width=""100%"">" _
            '        & CR  & vbTab & "<tr>" _
            '        & CR  & vbTab & vbTab & "<td style=""width:100%;vertical-align:top;text-align:left;"">" & Content & "</td>" _
            '        & CR  & vbTab & vbTab & "<td style=""width:1px;""><img src=""/ccLib/images/spacer.gif"" width=""1"" height=""400""></td>" _
            '        & CR  & vbTab & "</tr>" _
            '        & CR  & "</table>" _
            '        & kmaIndent(ButtonBar) _
            '        & CR  & cp.core.main_GetFormInputHidden(RequestNameRefreshBlock, cp.core.main_GetFormSN) _
            '        & ""
            'Stream.Add( ButtonBar)
            'Stream.Add( cp.core.main_GetFormInputHidden(RequestNameRefreshBlock, cp.core.main_GetFormSN)
            '
            GetBody = "" _
                & cr & cp.core.main_GetUploadFormStart() _
                & kmaIndent(GetBody) _
                & cr & cp.core.main_GetUploadFormEnd
            '
            Exit Function
            '
            ' ----- Error Trap
            '
ErrorTrap:
            Call HandleTrapError("GetBody")
        End Function
        '
        '
        '
        Public Function GetEditRow(ByVal HTMLFieldString As String, ByVal Caption As String, Optional ByVal HelpMessage As String = "", Optional ByVal FieldRequired As Boolean = False, Optional ByVal AllowActiveEdit As Boolean = False, Optional ByVal ignore0 As String = "") As String
            On Error GoTo ErrorTrap
            '
            Dim FastString As New fastStringClass
            Dim Copy As String
            Dim FormInputName As String
            '
            ' Left Side
            '
            Copy = Caption
            If Copy = "" Then
                Copy = "&nbsp;"
            End If
            GetEditRow = "<tr><td class=""ccEditCaptionCon""><nobr>" & Copy & "<img alt=""space"" src=""/ccLib/images/spacer.gif"" width=1 height=15 >"
            'GetEditRow = "<tr><td class=""ccAdminEditCaption""><nobr>" & Copy & "<img alt=""space"" src=""/ccLib/images/spacer.gif"" width=1 height=15 >"
            'If cp.core.main_VisitProperty_AllowHelpIcon Then
            '    'If HelpMessage <> "" Then
            '        GetEditRow = GetEditRow & "&nbsp;" & cp.core.main_GetHelpLinkEditable(0, Caption, HelpMessage, FormInputName)
            '    'Else
            '    '    GetEditRow = GetEditRow & "&nbsp;<img alt=""space"" src=""/ccLib/images/spacer.gif"" " & IconWidthHeight & ">"
            '    'End If
            'End If
            GetEditRow = GetEditRow & "</nobr></td>"
            '
            ' Right Side
            '
            Copy = HTMLFieldString
            If Copy = "" Then
                Copy = "&nbsp;"
            End If
            Copy = "<div class=""ccEditorCon"">" & Copy & "</div>"
            'If HelpMessage <> "" Then
            Copy = Copy & "<div class=""ccEditorHelpCon""><div class=""closed"">" & HelpMessage & "</div></div>"
            'Copy = Copy & "<div style=""padding:10px;white-space:normal"">" & HelpMessage & "</div>"
            'End If
            GetEditRow = GetEditRow & "<td class=""ccEditFieldCon"">" & Copy & "</td></tr>"
            'GetEditRow = GetEditRow & "<td class=""ccAdminEditField"">" & Copy & "</td></tr>"
            '
            Exit Function
            '
ErrorTrap:
            Call HandleTrapError("GetEditRow")
        End Function
        '
        '
        '
        Public Function GetEditRowWithHelpEdit(ByVal HTMLFieldString As String, ByVal Caption As String, Optional ByVal HelpMessage As String = "", Optional ByVal FieldRequired As Boolean = False, Optional ByVal AllowActiveEdit As Boolean = False, Optional ByVal ignore0 As String = "") As String
            On Error GoTo ErrorTrap
            '
            Dim FastString As New fastStringClass
            Dim Copy As String
            Dim FormInputName As String
            '
            Copy = Caption
            If Copy = "" Then
                Copy = "&nbsp;"
            End If
            GetEditRowWithHelpEdit = "<tr><td class=""ccAdminEditCaption""><nobr>" & Copy
            If cp.core.main_VisitProperty_AllowHelpIcon Then
                'If HelpMessage <> "" Then
                'GetEditRowWithHelpEdit = GetEditRowWithHelpEdit & "&nbsp;" & cp.core.main_GetHelpLinkEditable(0, Caption, HelpMessage, FormInputName)
                'Else
                '    GetEditRowWithHelpEdit = GetEditRowWithHelpEdit & "&nbsp;<img alt=""space"" src=""/ccLib/images/spacer.gif"" " & IconWidthHeight & ">"
                'End If
            End If
            GetEditRowWithHelpEdit = GetEditRowWithHelpEdit & "<img alt=""space"" src=""/ccLib/images/spacer.gif"" width=1 height=15 ></nobr></td>"
            Copy = HTMLFieldString
            If Copy = "" Then
                Copy = "&nbsp;"
            End If
            GetEditRowWithHelpEdit = GetEditRowWithHelpEdit & "<td class=""ccAdminEditField"">" & Copy & "</td></tr>"
            '
            Exit Function
            '
ErrorTrap:
            Call HandleTrapError("GetEditRowWithHelpEdit")
        End Function
        '
        '
        '
        Public Function GetEditSubheadRow(ByVal Caption As String) As String
            On Error GoTo ErrorTrap
            '
            GetEditSubheadRow = "<tr><td colspan=2 class=""ccAdminEditSubHeader"">" & Caption & "</td></tr>"
            'GetEditSubheadRow = "<tr><td colspan=2 class=""ccPanel3D ccAdminEditSubHeader""><b>" & Caption & "</b></td></tr>"
            '
            Exit Function
            '
ErrorTrap:
            Call HandleTrapError("GetEditSubheadRow")
        End Function
        '
        '========================================================================
        ' GetEditPanel
        '
        '   An edit panel is a section of an admin page, under a subhead.
        '   When in tab mode, the subhead is blocked, and the panel is assumed to
        '   go in its own tab windows
        '========================================================================
        '
        Public Function GetEditPanel(ByVal AllowHeading As Boolean, ByVal PanelHeading As String, ByVal PanelDescription As String, ByVal PanelBody As String) As String
            On Error GoTo ErrorTrap
            '
            Dim FastString As New fastStringClass
            '
            GetEditPanel = GetEditPanel & "<div class=""ccPanel3DReverse ccAdminEditBody"">"
            '
            ' ----- Subhead
            '
            If AllowHeading And (PanelHeading <> "") Then
                GetEditPanel = GetEditPanel & "<div class=""ccAdminEditHeading"">" & PanelHeading & "</div>"
            End If
            '
            ' ----- Description
            '
            If PanelDescription <> "" Then
                GetEditPanel = GetEditPanel & "<div class=""ccAdminEditDescription"">" & PanelDescription & "</div>"
            End If
            '
            GetEditPanel = GetEditPanel & PanelBody & "</div>"
            '
            ' ----- finish up panel
            '
            'EditSectionPanelCount = EditSectionPanelCount + 1
            '
            Exit Function
            '
ErrorTrap:
            Call HandleTrapError("GetEditPanel")
        End Function
        '
        '========================================================================
        ' Edit Table Open
        '========================================================================
        '
        Public ReadOnly Property EditTableOpen() As String
            Get
                EditTableOpen = "<table border=0 cellpadding=3 cellspacing=0 width=""100%"">"
            End Get
        End Property
        '
        '========================================================================
        ' Edit Table Close
        '========================================================================
        '
        Public ReadOnly Property EditTableClose() As String
            Get
                EditTableClose = "<tr>" _
                    & "<td width=20%><img alt=""space"" src=""/ccLib/images/spacer.gif"" width=""100%"" height=1 ></td>" _
                    & "<td width=80%><img alt=""space"" src=""/ccLib/images/spacer.gif"" width=""100%"" height=1 ></td>" _
                    & "</tr>" _
                    & "</table>"
            End Get

        End Property
        '
        '==========================================================================================
        '   Report Cell
        '==========================================================================================
        '
        Private Function GetReport_Cell(ByVal Copy As String, ByVal Align As String, ByVal Columns As Integer, ByVal RowPointer As Integer) As String
            Dim iAlign As String
            Dim Style As String
            Dim CellCopy As String
            '
            iAlign = encodeEmptyText(Align, "left")
            '
            If (RowPointer Mod 2) > 0 Then
                Style = "ccAdminListRowOdd"
            Else
                Style = "ccAdminListRowEven"
            End If
            '
            GetReport_Cell = vbCrLf & "<td valign=""middle"" align=""" & iAlign & """ class=""" & Style & """"
            If Columns > 1 Then
                GetReport_Cell = GetReport_Cell & " colspan=""" & Columns & """"
            End If
            '
            CellCopy = Copy
            If CellCopy = "" Then
                CellCopy = "&nbsp;"
            End If
            GetReport_Cell = GetReport_Cell & "><span class=""ccSmall"">" & CellCopy & "</span></td>"
        End Function
        '
        '==========================================================================================
        '   Report Cell Header
        '       ColumnPtr       :   0 based column number
        '       Title
        '       Width           :   if just a number, assumed to be px in style and an image is added
        '                       :   if 00px, image added with the numberic part
        '                       :   if not number, added to style as is
        '       align           :   style value
        '       ClassStyle      :   class
        '       RQS
        '       SortingState
        '==========================================================================================
        '
        Private Function GetReport_CellHeader(ByVal ColumnPtr As Integer, ByVal Title As String, ByVal Width As String, ByVal Align As String, ByVal ClassStyle As String, ByVal RefreshQueryString As String, ByVal SortingState As SortingStateEnum) As String
            '
            ' See new GetReportOrderBy for the method to setup sorting links
            '
            On Error GoTo ErrorTrap
            '
            Dim Style As String
            Dim Copy As String
            Dim QS As String
            Dim WidthTest As Integer
            Dim LinkTitle As String
            '
            If Title = "" Then
                Copy = "&nbsp;"
            Else
                Copy = Replace(Title, " ", "&nbsp;")
                'Copy = "<nobr>" & Title & "</nobr>"
            End If
            Style = "VERTICAL-ALIGN:bottom;"
            If Align = "" Then
            Else
                Style = Style & "TEXT-ALIGN:" & Align & ";"
            End If
            '
            Select Case SortingState
                Case SortingStateEnum.SortableNotSet
                    QS = ModifyQueryString(RefreshQueryString, "ColSort", CStr(SortingStateEnum.SortableSetAZ), True)
                    QS = ModifyQueryString(QS, "ColPtr", CStr(ColumnPtr), True)
                    Copy = "<a href=""?" & QS & """ title=""Sort A-Z"" class=""ccAdminListCaption"">" & Copy & "</a>"
                Case SortingStateEnum.SortableSetza
                    QS = ModifyQueryString(RefreshQueryString, "ColSort", CStr(SortingStateEnum.SortableSetAZ), True)
                    QS = ModifyQueryString(QS, "ColPtr", CStr(ColumnPtr), True)
                    Copy = "<a href=""?" & QS & """ title=""Sort A-Z"" class=""ccAdminListCaption"">" & Copy & "<img src=""/ccLib/images/arrowup.gif"" width=8 height=8 border=0></a>"
                Case SortingStateEnum.SortableSetAZ
                    QS = ModifyQueryString(RefreshQueryString, "ColSort", CStr(SortingStateEnum.SortableSetza), True)
                    QS = ModifyQueryString(QS, "ColPtr", CStr(ColumnPtr), True)
                    Copy = "<a href=""?" & QS & """ title=""Sort Z-A"" class=""ccAdminListCaption"">" & Copy & "<img src=""/ccLib/images/arrowdown.gif"" width=8 height=8 border=0></a>"
            End Select
            '
            If Width <> "" Then
                WidthTest = EncodeInteger(Replace(Width, "px", "", , , vbTextCompare))
                If WidthTest <> 0 Then
                    Style = Style & "width:" & WidthTest & "px;"
                    Copy = Copy & "<img alt=""space"" src=""/ccLib/images/spacer.gif"" width=""" & WidthTest & """ height=1 border=0>"
                    'Copy = Copy & "<br><img alt=""space"" src=""/ccLib/images/spacer.gif"" width=""" & WidthTest & """ height=1 border=0>"
                Else
                    Style = Style & "width:" & Width & ";"
                End If
            End If
            '
            GetReport_CellHeader = vbCrLf & "<td style=""" & Style & """ class=""" & ClassStyle & """>" & Copy & "</td>"
            '
            Exit Function
            '
            ' ----- Error Trap
            '
ErrorTrap:
            Call HandleTrapError("GetReport_CellHeader")
        End Function
        '
        '=============================================================================
        '   Get Sort Column Ptr
        '
        '   returns the integer column ptr of the column last selected
        '=============================================================================
        '
        Public Function GetReportSortColumnPtr(ByVal DefaultSortColumnPtr As Integer) As Integer
            Dim VarText As String
            '
            VarText = cp.core.main_GetStreamText2("ColPtr")
            GetReportSortColumnPtr = EncodeInteger(VarText)
            If (GetReportSortColumnPtr = 0) And (VarText <> "0") Then
                GetReportSortColumnPtr = DefaultSortColumnPtr
            End If
        End Function
        '
        '=============================================================================
        '   Get Sort Column Type
        '
        '   returns the integer for the type of sorting last requested
        '       0 = nothing selected
        '       1 = sort A-Z
        '       2 = sort Z-A
        '
        '   Orderby is generated by the selection of headers captions by the user
        '   It is up to the calling program to call GetReportOrderBy to get the orderby and use it in the query to generate the cells
        '   This call returns a comma delimited list of integers representing the columns to sort
        '=============================================================================
        '
        Public Function GetReportSortType() As Integer
            Dim VarText As String
            '
            VarText = cp.core.main_GetStreamText2("ColPtr")
            If (EncodeInteger(VarText) <> 0) Or (VarText = "0") Then
                '
                ' A valid ColPtr was found
                '
                GetReportSortType = cp.core.main_GetStreamInteger2("ColSort")
            Else
                GetReportSortType = SortingStateEnum.SortableSetAZ
            End If
        End Function
        '
        '=============================================================================
        '   Translate the old GetReport to the new GetReport2
        '=============================================================================
        '
        Public Function GetReport(ByVal RowCount As Integer, ByVal ColCaption() As String, ByVal ColAlign() As String, ByVal ColWidth() As String, ByVal Cells(,) As String, ByVal PageSize As Integer, ByVal PageNumber As Integer, ByVal PreTableCopy As String, ByVal PostTableCopy As String, ByVal DataRowCount As Integer, ByVal ClassStyle As String) As String
            On Error GoTo ErrorTrap
            '
            Dim ColSortable() As Boolean
            Dim ColCnt As Integer
            Dim Ptr As Integer
            '
            ColCnt = UBound(Cells, 2)
            ReDim ColSortable(ColCnt)
            For Ptr = 0 To ColCnt - 1
                ColSortable(Ptr) = False
            Next
            '
            GetReport = GetReport2(RowCount, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, ClassStyle, ColSortable, 0)
            '
            Exit Function
            '
            ' ----- Error Trap
            '
ErrorTrap:
            Call HandleTrapError("GetReport")
        End Function

        '
        '=============================================================================
        '   Report
        '
        '   This is a list report that you have to fill in all the cells and pass them in arrays.
        '   The column headers are always assumed to include the orderby options. they are linked. To get the correct orderby, the calling program
        '   has to call GetReport2orderby
        '
        '
        '=============================================================================
        '
        Public Function GetReport2(ByVal RowCount As Integer, ByVal ColCaption() As String, ByVal ColAlign() As String, ByVal ColWidth() As String, ByVal Cells As String(,), ByVal PageSize As Integer, ByVal PageNumber As Integer, ByVal PreTableCopy As String, ByVal PostTableCopy As String, ByVal DataRowCount As Integer, ByVal ClassStyle As String, ByVal ColSortable() As Boolean, ByVal DefaultSortColumnPtr As Integer) As String
            On Error GoTo ErrorTrap
            '
            Dim RQS As String
            Dim SortMethod As Integer
            'Dim ButtonBar As String
            Dim RowBAse As Integer
            Dim Content As New fastStringClass
            Dim Stream As New fastStringClass
            Dim ColumnCount As Integer
            Dim ColumnPtr As Integer
            Dim ColumnWidth As String
            Dim RowPointer As Integer
            Dim WorkingQS As String
            '
            Dim PageCount As Integer
            Dim PagePointer As Integer
            Dim LinkCount As Integer
            Dim ReportPageNumber As Integer
            Dim ReportPageSize As Integer
            Dim iClassStyle As String
            Dim SortColPtr As Integer
            Dim SortColType As Integer
            '
            ReportPageNumber = PageNumber
            If ReportPageNumber = 0 Then
                ReportPageNumber = 1
            End If
            ReportPageSize = PageSize
            If ReportPageSize < 1 Then
                ReportPageSize = 50
            End If
            '
            iClassStyle = ClassStyle
            If iClassStyle = "" Then
                iClassStyle = "ccPanel"
            End If
            'If IsArray(Cells) Then
            ColumnCount = UBound(Cells, 2)
            'End If
            RQS = cp.core.main_RefreshQueryString
            '
            SortColPtr = GetReportSortColumnPtr(DefaultSortColumnPtr)
            SortColType = GetReportSortType()
            '
            ' ----- Start the table
            '
            Call Content.Add(StartTable(3, 1, 0))
            '
            ' ----- Header
            '
            Call Content.Add(vbCrLf & "<tr>")
            Call Content.Add(GetReport_CellHeader(0, "Row", "50", "Right", "ccAdminListCaption", RQS, SortingStateEnum.NotSortable))
            For ColumnPtr = 0 To ColumnCount - 1
                ColumnWidth = ColWidth(ColumnPtr)
                If Not ColSortable(ColumnPtr) Then
                    '
                    ' not sortable column
                    '
                    Call Content.Add(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortingStateEnum.NotSortable))
                ElseIf ColumnPtr = SortColPtr Then
                    '
                    ' This is the current sort column
                    '
                    Call Content.Add(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortColType))
                Else
                    '
                    ' Column is sortable, but not selected
                    '
                    Call Content.Add(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortingStateEnum.SortableNotSet))
                End If

                'If ColumnPtr = SortColPtr Then
                '    '
                '    ' This column is currently the active sort
                '    '
                '    Call Content.Add(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortColType))
                'Else
                '    Call Content.Add(GetReport_CellHeader(ColumnPtr, ColCaption(ColumnPtr), ColumnWidth, ColAlign(ColumnPtr), "ccAdminListCaption", RQS, SortingStateEnum.SortableNotSet))
                'End If
            Next
            Call Content.Add(vbCrLf & "</tr>")
            '
            ' ----- Data
            '
            If RowCount = 0 Then
                Call Content.Add(vbCrLf & "<tr>")
                Call Content.Add(GetReport_Cell(RowBAse + RowPointer, "right", 1, RowPointer))
                Call Content.Add(GetReport_Cell("-- End --", "left", ColumnCount, 0))
                Call Content.Add(vbCrLf & "</tr>")
            Else
                RowBAse = (ReportPageSize * (ReportPageNumber - 1)) + 1
                For RowPointer = 0 To RowCount - 1
                    Call Content.Add(vbCrLf & "<tr>")
                    Call Content.Add(GetReport_Cell(RowBAse + RowPointer, "right", 1, RowPointer))
                    For ColumnPtr = 0 To ColumnCount - 1
                        Call Content.Add(GetReport_Cell(Cells(RowPointer, ColumnPtr), ColAlign(ColumnPtr), 1, RowPointer))
                    Next
                    Call Content.Add(vbCrLf & "</tr>")
                Next
            End If
            '
            ' ----- End Table
            '
            Call Content.Add(kmaEndTable)
            GetReport2 = GetReport2 & Content.Text
            '
            ' ----- Post Table copy
            '
            If (DataRowCount / ReportPageSize) <> Int((DataRowCount / ReportPageSize)) Then
                PageCount = (DataRowCount / ReportPageSize) + 0.5
            Else
                PageCount = (DataRowCount / ReportPageSize)
            End If
            If PageCount > 1 Then
                GetReport2 = GetReport2 & "<br>Page " & ReportPageNumber & " (Row " & (RowBAse) & " of " & DataRowCount & ")"
                If PageCount > 20 Then
                    PagePointer = ReportPageNumber - 10
                End If
                If PagePointer < 1 Then
                    PagePointer = 1
                End If
                If PageCount > 1 Then
                    GetReport2 = GetReport2 & "<br>Go to Page "
                    If PagePointer <> 1 Then
                        WorkingQS = cp.core.main_RefreshQueryString
                        WorkingQS = ModifyQueryString(WorkingQS, "GotoPage", "1", True)
                        GetReport2 = GetReport2 & "<a href=""" & cp.core.main_ServerPage & "?" & WorkingQS & """>1</A>...&nbsp;"
                    End If
                    WorkingQS = cp.core.main_RefreshQueryString
                    WorkingQS = ModifyQueryString(WorkingQS, RequestNamePageSize, CStr(ReportPageSize), True)
                    Do While (PagePointer <= PageCount) And (LinkCount < 20)
                        If PagePointer = ReportPageNumber Then
                            GetReport2 = GetReport2 & PagePointer & "&nbsp;"
                        Else
                            WorkingQS = ModifyQueryString(WorkingQS, RequestNamePageNumber, CStr(PagePointer), True)
                            GetReport2 = GetReport2 & "<a href=""" & cp.core.main_ServerPage & "?" & WorkingQS & """>" & PagePointer & "</A>&nbsp;"
                        End If
                        PagePointer = PagePointer + 1
                        LinkCount = LinkCount + 1
                    Loop
                    If PagePointer < PageCount Then
                        WorkingQS = ModifyQueryString(WorkingQS, RequestNamePageNumber, CStr(PageCount), True)
                        GetReport2 = GetReport2 & "...<a href=""" & cp.core.main_ServerPage & "?" & WorkingQS & """>" & PageCount & "</A>&nbsp;"
                    End If
                    If ReportPageNumber < PageCount Then
                        WorkingQS = ModifyQueryString(WorkingQS, RequestNamePageNumber, CStr(ReportPageNumber + 1), True)
                        GetReport2 = GetReport2 & "...<a href=""" & cp.core.main_ServerPage & "?" & WorkingQS & """>next</A>&nbsp;"
                    End If
                    GetReport2 = GetReport2 & "<br>&nbsp;"
                End If
            End If
            '
            GetReport2 = "" _
                & PreTableCopy _
                & "<table border=0 cellpadding=0 cellspacing=0 width=""100%""><tr><td style=""padding:10px;"">" _
                & GetReport2 _
                & "</td></tr></table>" _
                & PostTableCopy _
                & ""
            Exit Function
            '
            ' ----- Error Trap
            '
ErrorTrap:
            Call HandleTrapError("GetReport2")
        End Function
        '
        '========================================================================
        ' Get the Normal Edit Button Bar String
        '
        '   used on Normal Edit and others
        '========================================================================
        '
        Public Function GetEditButtonBar(ByVal MenuDepth As Integer, ByVal AllowDelete As Boolean, ByVal AllowCancel As Boolean, ByVal allowSave As Boolean, ByVal AllowSpellCheck As Boolean, ByVal AllowPublish As Boolean, ByVal AllowAbort As Boolean, ByVal AllowSubmit As Boolean, ByVal AllowApprove As Boolean) As String
            GetEditButtonBar = GetEditButtonBar2(MenuDepth, AllowDelete, AllowCancel, allowSave, AllowSpellCheck, AllowPublish, AllowAbort, AllowSubmit, AllowApprove, False, False, False, False, False, False, False)
        End Function
        '
        '========================================================================
        ' Get the Normal Edit Button Bar String
        '
        '   used on Normal Edit and others
        '========================================================================
        '
        Public Function GetFormBodyAdminOnly() As String
            GetFormBodyAdminOnly = "<div class=""ccError"" style=""margin:10px;padding:10px;background-color:white;"">This page requires administrator permissions.</div>"
        End Function
        '
        '========================================================================
        '   Builds a single entry in the ReportFilter at the bottom of the page
        '========================================================================
        '
        Public Function GetReportFilterRow(ByVal FormInput As String, ByVal Caption As String) As String
            '
            GetReportFilterRow = "" _
                & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" _
                & "<tr><td width=""200"" align=""right"" class=RowInput>" & FormInput & "</td>" _
                & "<td width=""100%"" align=""left"" class=RowCaption>&nbsp;" & Caption & "</td></tr>" _
                & "</table>"
            '
        End Function
        '
        '=============================================================================
        '   Builds the panel around the filters at the bottom of the report
        '=============================================================================
        '
        Public Function GetReportFilter(ByVal Title As String, ByVal Body As String) As String
            '
            GetReportFilter = GetReportFilter & "<div class=""ccReportFilter"">"
            GetReportFilter = GetReportFilter & "<div class=""Title"">" & Title & "</div>"
            GetReportFilter = GetReportFilter & "<div class=""Body"">" & Body & "</div>"
            GetReportFilter = GetReportFilter & "</div>"
            '
        End Function



        Public Sub New(cp As CPBaseClass)
            MyBase.New()
            Me.cp = cp
        End Sub
    End Class
End Namespace
