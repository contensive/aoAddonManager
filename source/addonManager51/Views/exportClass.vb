
Imports Contensive.Addons.AddonManager51.Models
Imports Contensive.BaseClasses
Imports Contensive.Models.Db
Imports ICSharpCode.SharpZipLib

Namespace Contensive.Addons.AddonManager51
    Public Class exportClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        ''' <summary>
        ''' export collection
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Try
                Dim Button As String = CP.Doc.GetText(RequestNameButton)
                If Button = ButtonCancel Then
                    '
                    ' ----- redirect back to the root
                    Call CP.Response.Redirect(CP.Site.GetText("adminUrl"))
                    Return String.Empty
                End If
                '
                ' -- create form
                Dim form As New adminFramework.formNameValueRowsClass With {
                    .title = "Export Collection",
                    .body = CP.Html.p("Use this tool to create an Add-on Collection zip file that can be used to install a collection on another site.")
                }
                If Not CP.User.IsAdmin() Then
                    '
                    ' -- Put up error message
                    form.body &= CP.Html.p("You must be an administrator to use this tool.")
                    Return form.getHtml(CP)
                End If
                If (CP.Site.GetText("buildVersion") < CP.Version) Then
                    '
                    ' -- database needs to be upgraded
                    form.description &= CP.Html.p("Warning: The site's Database needs to be upgraded. You should do this before installing addons.")
                End If
                '
                ' -- Upload tool
                If (Button = ButtonOK) Then
                    '
                    ' -- export
                    Dim CollectionID As Integer = CP.Doc.GetInteger(RequestNameCollectionID)
                    Dim addonCollection As AddonCollectionModel = AddonCollectionModel.create(Of AddonCollectionModel)(CP, CollectionID)
                    If (addonCollection Is Nothing) Then
                        '
                        ' -- collection not found
                        Call CP.UserError.Add("The collection file you selected could not be found. Please select another.")
                    Else
                        '
                        ' -- build collection zip file and return file
                        Dim CollectionFilename As String = ""
                        Dim userError As String = ""
                        Try
                            ''
                            ' -- attempt new method
                            CP.Addon.ExportCollection(CollectionID, CollectionFilename, userError)
                        Catch ex As Exception
                            '
                            ' -- missing method, use the internal method
                            CollectionFilename = LegacyExportController.createCollectionZip_returnCdnPathFilename(CP, CollectionID)
                        End Try
                        If Not CP.UserError.OK Then
                            '
                            ' -- errors during export
                            form.body = CP.Html.div(CP.Html.p("ERRORS during export: ") & CP.Html.ul(CP.UserError.GetList()))
                        Else
                            '
                            ' -- success
                            form.body &= CP.Html.p("Export Successful")
                        End If
                        form.body &= CP.Html.p("Click <a href=""" & CP.Site.FilePath & Replace(CollectionFilename, "\", "/") & """>here</a> to download the collection file.</p>")
                    End If
                End If
                '
                ' Get Form
                form.addRow()
                form.rowName = "Add-on Collection"
                form.rowValue = CP.Html.SelectContent(RequestNameCollectionID, "0", "Add-on Collections", "", "Select Collection To Export", "form-control")
                form.addFormHidden("UploadCount", "1")
                form.addFormButton(ButtonOK)
                form.addFormButton(ButtonCancel)
                Return form.getHtml(CP)
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
                Return String.Empty
            End Try
        End Function
    End Class
End Namespace
