
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports Contensive.addonManager

Namespace Contensive.addonManager
    Module genericController
        '
        Public Function kmaIndent(source As String) As String
            Return source
        End Function
        '
        '
        '========================================================================
        ' ----- Get an XML nodes attribute based on its name
        '========================================================================
        '
        Public Function GetXMLAttribute(found As Boolean, Node As Xml.XmlNode, Name As String, DefaultIfNotFound As String) As String
            Dim result As String = ""
            Try
                Dim REsultNode As Xml.XmlNode
                Dim UcaseName As String
                '
                found = False
                REsultNode = Node.Attributes.GetNamedItem(Name)
                If (REsultNode Is Nothing) Then
                    UcaseName = UCase(Name)
                    For Each kvp As KeyValuePair(Of String, System.Xml.XmlAttribute) In Node.Attributes

                        If kvp.Value.Name.ToUpper() = UcaseName Then
                            result = kvp.Value.Value
                            found = True
                            Exit For
                        End If
                    Next
                Else
                    result = REsultNode.Value
                    found = True
                End If
                If Not found Then
                    result = DefaultIfNotFound
                End If
            Catch ex As Exception
                Throw
            End Try
            Return result
        End Function

        '
        '
        '
        Public Sub GetForm_AddonManager_DeleteNavigatorBranch(cp As CPBaseClass, EntryName As String, EntryParentID As Integer)
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
                        Call GetForm_AddonManager_DeleteNavigatorBranch(cp, cs.GetText("name"), EntryID)
                        Call cs.GoNext()
                    Loop
                    Call cs.Close()
                    Call cp.Content.Delete("Navigator Entries", "id=" & EntryID)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
        End Sub


        '
        '
        '
        Public Function GetParentIDFromNameSpace(cp As CPBaseClass, ContentName As String, NameSpacex As String) As Integer
            Dim result As Integer = 0
            Try
                Dim ParentNameSpace As String
                Dim ParentName As String
                Dim ParentID As Integer
                Dim Pos As Integer
                Dim cs As CPCSBaseClass = cp.CSNew
                '
                If NameSpacex <> "" Then
                    Pos = InStr(1, NameSpacex, ".")
                    If Pos = 0 Then
                        ParentName = NameSpacex
                        ParentNameSpace = ""
                    Else
                        ParentName = Mid(NameSpacex, Pos + 1)
                        ParentNameSpace = Mid(NameSpacex, 1, Pos - 1)
                    End If
                    If ParentNameSpace = "" Then
                        If cs.Open(ContentName, "(name=" & cp.Db.EncodeSQLText(ParentName) & ")and((parentid is null)or(parentid=0))", "ID", , "ID") Then
                            result = cs.GetInteger("ID")
                        End If
                        Call cs.Close()
                    Else
                        ParentID = GetParentIDFromNameSpace(cp, ContentName, ParentNameSpace)
                        If cs.Open(ContentName, "(name=" & cp.Db.EncodeSQLText(ParentName) & ")and(parentid=" & ParentID & ")", "ID",  , "ID") Then
                            result = cs.GetInteger("ID")
                        End If
                        Call cs.Close()
                    End If
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
        '
        '
        '
        Public Function getLayout(cp As CPBaseClass, layoutGuid As String, DefaultName As String) As String
            Dim result As String = ""
            Try
                Dim cs As CPCSBaseClass = cp.CSNew
                '
                cs.Open("layouts", "ccguid=" & cp.Db.EncodeSQLText(layoutGuid))
                If Not cs.OK Then
                    Call cs.Close()
                    If cs.Insert("layouts") Then
                        Call cs.SetField("ccguid", layoutGuid)
                        Call cs.SetField("name", DefaultName)
                        Call cs.SetField("layout", "<!-- layout record created " & Now & " -->")
                    End If
                End If
                If cs.OK Then
                    result = cs.GetText("layout")
                End If
                Call cs.Close()
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return result
        End Function
    End Module
End Namespace

