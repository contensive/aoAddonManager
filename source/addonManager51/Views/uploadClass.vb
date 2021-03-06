Imports Contensive.BaseClasses

Namespace Contensive.Addons.AddonManager51
    '
    Public Class UploadClass
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
            Try
                Me.cp = CP
                Return getUpload()
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Throw
            End Try
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
                    Dim form As New PortalFramework.FormNameValueRowsClass With {
                        .title = "Upload Collection",
                        .body = cp.Html.p("Use this form to upload an add-on collection. If the GUID of the add-on matches one already installed on this server, it will be updated. If the GUID is new, it will be added."),
                        .description = cp.Html.p("Upload a collection zip file to install the collection on this site. ")
                    }
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
                                '
                                ' -- use v5 method
                                Dim ErrorMessage As String = ""
                                If (Not InstallController.installCollectionFromLibrary(cp, "{8DAABAE6-8E45-4CEE-A42C-B02D180E799B}", ErrorMessage)) Then
                                    form.description &= cp.Html.p("ERROR: " & ErrorMessage)
                                End If
                            End If
                            Dim uploadFilename As String = cp.Doc.GetText(rnUploadCollectionFile)
                            If (Not String.IsNullOrEmpty(uploadFilename)) Then
                                '
                                ' -- version 5.0, separate class so this project can be built with contensive 5.0 reference, but run against contensive 4.1
                                Dim ErrorMessage As String = ""
                                Dim installDependencies As Boolean = Not cp.Doc.GetBoolean(rnBlockDependencies)
                                If (InstallController.installCollectionFromUpload(cp, rnUploadCollectionFile, ErrorMessage)) Then
                                    If (Not String.IsNullOrEmpty(ErrorMessage)) Then
                                        '
                                        ' -- install successful, but a problem
                                        form.body &= cp.Html.p("Installation completed with the follow message [" & ErrorMessage & "].")
                                    Else
                                        '
                                        ' -- install failed
                                        form.body &= cp.Html.p("Installation successful.")
                                    End If
                                Else
                                    form.body &= cp.Html.p("Error installing collection files, ERROR: " & ErrorMessage)
                                End If
                            End If
                        End If
                        '
                        ' Get Form
                        form.addRow()
                        form.rowName = "Collection Zip File"
                        form.rowValue = cp.Html.InputFile(rnUploadCollectionFile)
                        form.addRow()
                        form.rowName = "Block Dependencies"
                        form.rowValue = cp.Html.CheckBox(rnBlockDependencies)
                        form.addFormHidden("UploadCount", "1")
                        form.addFormButton(ButtonOK)
                        form.addFormButton(ButtonCancel)
                    End If
                    returnResult = form.getHtml(cp)
                End If
                Return returnResult
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
                Throw
            End Try
        End Function
    End Class
End Namespace
