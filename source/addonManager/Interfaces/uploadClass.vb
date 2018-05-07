
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace Contensive.Addons.AddonManager
    '
    Public Class uploadClass
        Inherits AddonBaseClass
        '
        ' -- injected objects -- do not dispose
        Private cp As CPBaseClass
        ' 
        ' -- class scope
        Private Const RequestNameButton As String = "button"
        Private Const ButtonCancel As String = " Cancel "
        Private Const ButtonOK As String = " OK "
        '
        '====================================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Me.cp = CP
                returnHtml = getUpload()
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
            End Try
            Return returnHtml
        End Function
        '
        '====================================================================================================
        '
        Private Function getUpload() As String
            Dim returnResult As String = ""
            Try
                Dim Button As String = cp.Doc.GetText(RequestNameButton)
                If Button = ButtonCancel Then
                    '
                    ' ----- redirect back to the root
                    Call cp.Response.Redirect(cp.Site.GetText("adminUrl"))
                Else
                    '
                    ' -- create form
                    Dim form As New adminFramework.formNameValueRowsClass
                    form.title = "Upload Collection"
                    form.body = cp.Html.p("Use this form to upload an add-on collection. If the GUID of the add-on matches one already installed on this server, it will be updated. If the GUID is new, it will be added.")
                    form.description = cp.Html.p("Upload a collection zip file to install the collection on this site. ")
                    If Not cp.User.IsAdmin() Then
                        '
                        ' -- Put up error message
                        form.body &= cp.Html.p("You must be an administrator to use this tool.")
                    Else
                        If (cp.Site.GetText("buildVersion") < cp.Version) Then
                            '
                            ' -- database needs to be upgraded
                            form.description &= cp.Html.p("WARNING: The site database needs to be upgraded. You should do this before installing addon collections.")
                        End If
                        '
                        ' -- Upload tool
                        If (Button = ButtonOK) Then
                            '
                            ' -- handle upload
                            If cp.User.IsDeveloper And cp.Doc.GetBoolean("InstallCore") Then
                                '
                                ' -- Reinstall core collection
                                Call cp.Content.Delete("Add-on Collections", "ccguid='{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}'")
                                If (cp.Version < "5") Then
                                    '
                                    ' -- use v4 methods, [addonManager-InstallFromLibraryC41]
                                    cp.Doc.SetProperty("collectionGuid", "{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}")
                                    If (Not cp.Utils.EncodeBoolean(cp.Utils.ExecuteAddon(legacyMethodInstallFromLibraryC41))) Then
                                        form.description &= cp.Html.p("ERROR: There was an unknown error installing the collection. Details are in the server log file(s).")
                                    End If
                                Else
                                    '
                                    ' -- use v5 method
                                    Dim ErrorMessage As String = ""
                                    If (Not v5InstallController.installCollectionFromLibrary(cp, "{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}", ErrorMessage)) Then
                                        form.description &= cp.Html.p("ERROR: " & ErrorMessage)
                                    End If
                                End If
                            End If
                            Dim uploadFilename As String = cp.Doc.GetText(rnUploadCollectionFile)
                            If (Not String.IsNullOrEmpty(uploadFilename)) Then
                                If (cp.Version < "5") Then
                                    '
                                    ' -- version 4.1
                                    Dim InstallPath As String = "CollectionUpload" & cp.Utils.CreateGuid().Replace("{", "").Replace("-", "").Replace("}", "") & "\"
                                    cp.Html.ProcessInputFile(rnUploadCollectionFile, InstallPath)
                                    cp.Doc.SetProperty("physicalInstallPath", cp.Site.PhysicalFilePath & InstallPath)
                                    If cp.Utils.EncodeBoolean(cp.Utils.ExecuteAddon(legacyMethodInstallFromPhysicalInstallPathC41)) Then
                                        form.body &= cp.Html.p("Installed collection files.")
                                    Else
                                        form.body &= cp.Html.p("Error installing collection files.")
                                    End If
                                Else
                                    '
                                    ' -- version 5.0, separate class so this project can be built with contensive 5.0 reference, but run against contensive 4.1
                                    Dim ErrorMessage As String = ""
                                    If (v5InstallController.installCollectionFromUpload(cp, rnUploadCollectionFile, ErrorMessage)) Then
                                        form.body &= cp.Html.p("Installed collection files.")
                                    Else
                                        form.body &= cp.Html.p("Error installing collection files, ERROR: " & ErrorMessage)
                                    End If
                                End If
                            End If
                        End If
                        '
                        ' Get Form
                        form.addRow()
                        form.rowName = "Add-on Collection File"
                        form.rowValue = cp.Html.InputFile(rnUploadCollectionFile)
                        form.addFormHidden("UploadCount", "1")
                        form.addFormButton(ButtonOK)
                        form.addFormButton(ButtonCancel)
                    End If
                    returnResult = form.getHtml(cp)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return returnResult
        End Function
    End Class
End Namespace
