﻿
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports Contensive.addonManager

Namespace Contensive.addonManager
    Public Class v5InstallController
        '
        ' -- method provided here because these methods are not included in the c41 interface, so this call can only be created if v5 code
        Public Shared Function installCollectionFromLibrary(cp As CPBaseClass, collectionGuid As String, ByRef ErrorMessage As String) As Boolean
            Return cp.Addon.installCollectionFromLibrary(collectionGuid, ErrorMessage)
        End Function
        '
        ' -- method provided here because these methods are not included in the c41 interface, so this call can only be created if v5 code
        Public Shared Function installCollectionFromFolder(cp As CPBaseClass, privatePathFilename As String, ByRef ErrorMessage As String) As Boolean
            '
            cp.Utils.AppendLogFile("installCollectionFromFolder, privatePathFilename [" & privatePathFilename & "]")
            '
            Try
                Dim taskId As Integer = cp.Utils.installCollectionFromFile(privatePathFilename)
                ErrorMessage = ""
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return True
        End Function
        '
        ' -- method provided here because these methods are not included in the c41 interface, so this call can only be created if v5 code
        Public Shared Function installCollectionFromUpload(cp As CPBaseClass, requestName As String, ByRef ErrorMessage As String) As Boolean
            '
            cp.Utils.AppendLogFile("installCollectionFromUpload, requestName [" & requestName & "]")
            Try
                '
                Dim privatePath As String = "CollectionUpload" & cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "") & "\"

                Dim uploadFilename As String = ""
                If cp.privateFiles.saveUpload(requestName, privatePath, uploadFilename) Then
                    Return installCollectionFromFolder(cp, privatePath & uploadFilename, ErrorMessage)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return False
        End Function
    End Class
End Namespace
