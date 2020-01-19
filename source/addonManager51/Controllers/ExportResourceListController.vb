
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.Addons.AddonManager51.Models
Imports Contensive.BaseClasses
Imports ICSharpCode.SharpZipLib

Namespace Contensive.Addons.AddonManager51
    Public Class ExportResourceListController
        '
        '====================================================================================================
        '
        Public Shared Function getResourceList(cp As CPBaseClass, execFileList As String, CollectionGuid As String, tempPathFileList As List(Of String), tempExportPath As String) As String
            Try
                Dim nodeList As String = ""
                If execFileList <> "" Then
                    Dim LastChangeDate As Date
                    '
                    ' There are executable files to include in the collection
                    '   If installed, source path is collectionpath, if not installed, collectionpath will be empty
                    '   and file will be sourced right from addon path
                    '
                    Dim CollectionPath As String = ""
                    Call LegacyExportController.GetLocalCollectionArgs(cp, CollectionGuid, CollectionPath, LastChangeDate)
                    If CollectionPath <> "" Then
                        CollectionPath = CollectionPath & "\"
                    End If
                    Dim Files() As String = Split(execFileList, vbCrLf)
                    For Ptr As Integer = 0 To UBound(Files)
                        Dim PathFilename As String = Files(Ptr)
                        If PathFilename <> "" Then
                            PathFilename = Replace(PathFilename, "\", "/")
                            Dim Path As String = ""
                            Dim Filename As String = PathFilename
                            Dim Pos As Integer = InStrRev(PathFilename, "/")
                            If Pos > 0 Then
                                Filename = Mid(PathFilename, Pos + 1)
                                Path = Mid(PathFilename, 1, Pos - 1)
                            End If
                            Dim ManualFilename As String = ""
                            If LCase(Filename) <> LCase(ManualFilename) Then
                                Dim AddonPath As String = "addons\"
                                'AddFilename = AddonPath & CollectionPath & Filename
                                cp.PrivateFiles.Copy(AddonPath & CollectionPath & Filename, tempExportPath & Filename, cp.TempFiles)
                                If Not tempPathFileList.Contains(tempExportPath & Filename) Then
                                    tempPathFileList.Add(tempExportPath & Filename)
                                    nodeList = nodeList & vbCrLf & vbTab & "<Resource name=""" & System.Net.WebUtility.HtmlEncode(Filename) & """ type=""executable"" path=""" & System.Net.WebUtility.HtmlEncode(Path) & """ />"
                                End If
                            End If
                        End If
                    Next
                End If
                Return nodeList
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Return String.Empty
            End Try
        End Function
    End Class
End Namespace