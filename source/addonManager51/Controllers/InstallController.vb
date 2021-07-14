Imports Contensive.BaseClasses

Namespace Contensive.Addons.AddonManager51
    Public NotInheritable Class InstallController
        '
        ' -- method provided here because these methods are not included in the c41 interface, so this call can only be created if v5 code
        Public Shared Function installCollectionFromLibrary(cp As CPBaseClass, collectionGuid As String, ByRef ErrorMessage As String) As Boolean
            '
            cp.Utils.AppendLog("installCollectionFromLibrary, collectionGuid [" & collectionGuid & "]")
            '
            cp.Addon.InstallCollectionFromLibrary(collectionGuid, ErrorMessage)
            Return True
        End Function
        '
        ' -- method provided here because these methods are not included in the c41 interface, so this call can only be created if v5 code
        Public Shared Function installCollectionFromFolder(cp As CPBaseClass, privatePathFilename As String, ByRef ErrorMessage As String) As Boolean
            Try
                '
                cp.Utils.AppendLog("installCollectionFromFolder, privatePathFilename [" & privatePathFilename & "]")
                '
                Return cp.Addon.InstallCollectionFile(privatePathFilename, ErrorMessage)
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
        '
        ' -- method provided here because these methods are not included in the c41 interface, so this call can only be created if v5 code
        Public Shared Function installCollectionFromUpload(cp As CPBaseClass, requestName As String, ByRef ErrorMessage As String) As Boolean
            '
            cp.Utils.AppendLog("installCollectionFromUpload, requestName [" & requestName & "]")
            Try
                '
                Dim privatePath As String = "CollectionUpload" & cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "") & "\"
                Dim uploadFilename As String = ""
                Dim result As Boolean = False
                If cp.PrivateFiles.SaveUpload(requestName, privatePath, uploadFilename) Then
                    result = installCollectionFromFolder(cp, privatePath & uploadFilename, ErrorMessage)
                End If
                cp.PrivateFiles.DeleteFolder(privatePath)
                Return result
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
    End Class
End Namespace

