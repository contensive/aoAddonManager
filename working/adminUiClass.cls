VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "adminUiClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Public Enum SortingStateEnum
    NotSortable = 0
    SortableSetAZ = 1
    SortableSetza = 2
    SortableNotSet = 3
End Enum
'
'===========================================================================
'
'===========================================================================
'
Public Function GetTitleBar(main As Object, Title As String, Description As String) As String
    Dim worker As New cpCom.adminUiClass
    GetTitleBar = worker.GetTitleBar(main.getCmc, Title, Description)
End Function
'
'========================================================================
' Return a panel header with the header message reversed out of the left
'========================================================================
'
Public Function GetHeader(main As Object, HeaderMessage As String, Optional RightSideMessage As String) As String
    Dim worker As New cpCom.adminUiClass
    GetHeader = worker.GetHeader(main.getCmc, HeaderMessage, RightSideMessage)
End Function
'
'
'
Public Function GetButtonsFromList(main As Object, ButtonList As String, AllowDelete As Boolean, AllowAdd As Boolean, ButtonName As String) As String
    Dim worker As New cpCom.adminUiClass
    GetButtonsFromList = worker.GetButtonsFromList(main.getCmc, ButtonList, AllowDelete, AllowAdd, ButtonName)
End Function
'
'
'
Public Function GetButtonBar(LeftButtons As String, RightButtons As String) As String
    Dim worker As New cpCom.adminUiClass
    GetButtonBar = worker.GetButtonBar(LeftButtons, RightButtons)
End Function
'
'========================================================================
'
'========================================================================
'
Public Function GetBody(main As Object, Caption As String, ButtonListLeft As String, ButtonListRight As String, AllowAdd As Boolean, AllowDelete As Boolean, Description As String, ContentSummary As String, ContentPadding As Long, Content As String) As String
    Dim worker As New cpCom.adminUiClass
    GetBody = worker.GetBody(main.getCmc, Caption, ButtonListLeft, ButtonListRight, AllowAdd, AllowDelete, Description, ContentSummary, ContentPadding, Content)
End Function
'
'
'
Public Function GetEditRow(main As Object, HTMLFieldString As String, Caption As String, Optional HelpMessage As String, Optional FieldRequired As Boolean, Optional AllowActiveEdit As Boolean, Optional ignore0 As String) As String
    Dim worker As New cpCom.adminUiClass
    GetEditRow = worker.GetEditRow(main.getCmc, HTMLFieldString, Caption, HelpMessage, FieldRequired, AllowActiveEdit, ignore0)
End Function
'
'
'
Public Function GetEditRowWithHelpEdit(main As Object, HTMLFieldString As String, Caption As String, Optional HelpMessage As String, Optional FieldRequired As Boolean, Optional AllowActiveEdit As Boolean, Optional ignore0 As String) As String
    Dim worker As New cpCom.adminUiClass
    GetEditRowWithHelpEdit = worker.GetEditRowWithHelpEdit(main.getCmc, HTMLFieldString, Caption, HelpMessage, FieldRequired, AllowActiveEdit, ignore0)
End Function
'
'
'
Public Function GetEditSubheadRow(main As Object, Caption As String) As String
    Dim worker As New cpCom.adminUiClass
    GetEditSubheadRow = worker.GetEditSubheadRow(main.getCmc, Caption)
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
Public Function GetEditPanel(main As Object, AllowHeading As Boolean, PanelHeading As String, PanelDescription As String, PanelBody As String) As String
    Dim worker As New cpCom.adminUiClass
    GetEditPanel = worker.GetEditPanel(main.getCmc, AllowHeading, PanelHeading, PanelDescription, PanelBody)
End Function
'
'========================================================================
' Edit Table Open
'========================================================================
'
Public Property Get EditTableOpen() As String
    Dim worker As New cpCom.adminUiClass
    EditTableOpen = worker.EditTableOpen
End Property
'
'========================================================================
' Edit Table Close
'========================================================================
'
Public Property Get EditTableClose() As String
    Dim worker As New cpCom.adminUiClass
    EditTableClose = worker.EditTableClose
End Property
'
'=============================================================================
'   Get Sort Column Ptr
'
'   returns the integer column ptr of the column last selected
'=============================================================================
'
Public Function GetReportSortColumnPtr(main As Object, DefaultSortColumnPtr As Long) As Long
    Dim worker As New cpCom.adminUiClass
    GetReportSortColumnPtr = worker.GetReportSortColumnPtr(main.getCmc, DefaultSortColumnPtr)
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
Public Function GetReportSortType(main As Object) As Long
    Dim worker As New cpCom.adminUiClass
    GetReportSortType = worker.GetReportSortType(main.getCmc)
End Function
'
'=============================================================================
'   Translate the old GetReport to the new GetReport2
'=============================================================================
'
Public Function GetReport(main As Object, RowCount As Long, ColCaption() As String, ColAlign() As String, ColWidth() As String, Cells() As String, PageSize As Long, PageNumber As Long, PreTableCopy As String, PostTableCopy As String, DataRowCount As Long, ClassStyle As String) As String
    Dim worker As New cpCom.adminUiClass
    GetReport = worker.GetReport(main.getCmc, RowCount, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, ClassStyle)
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
Public Function GetReport2(main As Object, RowCount As Long, ColCaption() As String, ColAlign() As String, ColWidth() As String, Cells() As String, PageSize As Long, PageNumber As Long, PreTableCopy As String, PostTableCopy As String, DataRowCount As Long, ClassStyle As String, ColSortable() As Boolean, DefaultSortColumnPtr As Long) As String
    Dim worker As New cpCom.adminUiClass
    GetReport2 = worker.GetReport2(main.getCmc, RowCount, ColCaption, ColAlign, ColWidth, Cells, PageSize, PageNumber, PreTableCopy, PostTableCopy, DataRowCount, ClassStyle, ColSortable, DefaultSortColumnPtr)
End Function
'
'========================================================================
' Get the Normal Edit Button Bar String
'
'   used on Normal Edit and others
'========================================================================
'
Public Function GetEditButtonBar(main As Object, MenuDepth As Long, AllowDelete As Boolean, AllowCancel As Boolean, allowSave As Boolean, AllowSpellCheck As Boolean, AllowPublish As Boolean, AllowAbort As Boolean, AllowSubmit As Boolean, AllowApprove As Boolean) As String
    Dim worker As New cpCom.adminUiClass
    GetEditButtonBar = worker.GetEditButtonBar(main.getCmc, MenuDepth, AllowDelete, AllowCancel, allowSave, AllowSpellCheck, AllowPublish, AllowAbort, AllowSubmit, AllowApprove)
End Function
'
'========================================================================
' Get the Normal Edit Button Bar String
'
'   used on Normal Edit and others
'========================================================================
'
Public Function GetFormBodyAdminOnly(main As Object) As String
    Dim worker As New cpCom.adminUiClass
    GetFormBodyAdminOnly = worker.GetFormBodyAdminOnly(main.getCmc)
End Function
'
'========================================================================
'   Builds a single entry in the ReportFilter at the bottom of the page
'========================================================================
'
Public Function GetReportFilterRow(FormInput As String, Caption As String) As String
    Dim worker As New cpCom.adminUiClass
    GetReportFilterRow = worker.GetReportFilterRow(FormInput, Caption)
End Function
'
'=============================================================================
'   Builds the panel around the filters at the bottom of the report
'=============================================================================
'
Public Function GetReportFilter(Title As String, Body As String) As String
    Dim worker As New cpCom.adminUiClass
    GetReportFilter = worker.GetReportFilter(Title, Body)
End Function

