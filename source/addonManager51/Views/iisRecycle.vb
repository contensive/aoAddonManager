

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports ICSharpCode.SharpZipLib
Imports Microsoft.Web.Administration

Namespace Contensive.Addons.AddonManager51
    Public Class iisRecycleClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                'Using myProcess As New Process
                '    myProcess.StartInfo.UseShellExecute = False
                '    myProcess.StartInfo.FileName = "c:\windows\system32\inetsrv\appcmd recycle apppool """ & CP.Site.Name & """"
                '    myProcess.StartInfo.CreateNoWindow = True
                '    myProcess.Start()
                'End Using
                ' -- permission issue. user must be in administrators group. cant happen
                Using iisManager As New ServerManager()
                    Dim sites As SiteCollection = iisManager.Sites
                    For Each site As Site In sites
                        If (site.Name = System.Web.Hosting.HostingEnvironment.ApplicationHost.GetSiteName()) Then
                            iisManager.ApplicationPools(site.Applications("/").ApplicationPoolName).Recycle()
                            Exit For
                        End If
                    Next
                End Using
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
            End Try
            Return returnHtml
        End Function
    End Class
End Namespace
