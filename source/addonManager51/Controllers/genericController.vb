Imports System.IO
Imports Contensive.BaseClasses

Namespace Contensive.Addons.AddonManager51
    Module GenericController
        '
        '
        '
        Public Sub getForm_AddonManager_DeleteNavigatorBranch(cp As CPBaseClass, EntryName As String, EntryParentID As Integer)
            Try
                '
                Dim cs As CPCSBaseClass = cp.CSNew
                Dim EntryID As Integer
                '
                If EntryParentID = 0 Then
                    cs.Open("Navigator Entries", "(name=" & cp.Db.EncodeSQLText(EntryName) & ")and((parentID is null)or(parentid=0))")
                Else
                    cs.Open("Navigator Entries", "(name=" & cp.Db.EncodeSQLText(EntryName) & ")and(parentID=" & cp.Db.EncodeSQLNumber(EntryParentID) & ")")
                End If
                If cs.OK Then
                    EntryID = cs.GetInteger("ID")
                End If
                Call cs.Close()
                '
                If EntryID <> 0 Then
                    cs.Open("Navigator Entries", "(parentID=" & cp.Db.EncodeSQLNumber(EntryID) & ")")
                    Do While cs.OK
                        Call getForm_AddonManager_DeleteNavigatorBranch(cp, cs.GetText("name"), EntryID)
                        Call cs.GoNext()
                    Loop
                    Call cs.Close()
                    Call cp.Content.Delete("Navigator Entries", "id=" & EntryID)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Sub
        '
        '
        '
        Public Function getParentIDFromNameSpace(cp As CPBaseClass, ContentName As String, NameSpacex As String) As Integer
            Dim result As Integer = 0
            Try
                Dim ParentNameSpace As String
                Dim ParentName As String
                Dim ParentID As Integer
                Dim Pos As Integer
                Dim cs As CPCSBaseClass = cp.CSNew
                '
                If Not String.IsNullOrEmpty(NameSpacex) Then
                    Pos = InStr(1, NameSpacex, ".")
                    If Pos = 0 Then
                        ParentName = NameSpacex
                        ParentNameSpace = ""
                    Else
                        ParentName = Mid(NameSpacex, Pos + 1)
                        ParentNameSpace = Mid(NameSpacex, 1, Pos - 1)
                    End If
                    If String.IsNullOrEmpty(ParentNameSpace) Then
                        If cs.Open(ContentName, "(name=" & cp.Db.EncodeSQLText(ParentName) & ")and((parentid is null)or(parentid=0))", "ID", True, "ID") Then
                            result = cs.GetInteger("ID")
                        End If
                        Call cs.Close()
                    Else
                        ParentID = getParentIDFromNameSpace(cp, ContentName, ParentNameSpace)
                        If cs.Open(ContentName, "(name=" & cp.Db.EncodeSQLText(ParentName) & ")and(parentid=" & ParentID & ")", "ID", True, "ID") Then
                            result = cs.GetInteger("ID")
                        End If
                        Call cs.Close()
                    End If
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
            Return result
        End Function
    End Module
End Namespace

