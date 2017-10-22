
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
        Public Shared Function installCollectionFromFolder(cp As CPBaseClass, physicalFileFolder As String, ByRef ErrorMessage As String) As Boolean
            Dim taskId As Integer = cp.Utils.installCollectionFromFile(physicalFileFolder)
            Return True
        End Function
    End Class
End Namespace

