
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.Addons.AddonManager51.Models
Imports Contensive.BaseClasses
Imports ICSharpCode.SharpZipLib

Namespace Contensive.Addons.AddonManager51
    Public Class _BlankNodeListController
        '
        '====================================================================================================
        '
        Public Shared Function getNodeList(cp As CPBaseClass, execFileList As String) As String
            Try
                Return String.Empty
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Return String.Empty
            End Try
        End Function
    End Class
End Namespace