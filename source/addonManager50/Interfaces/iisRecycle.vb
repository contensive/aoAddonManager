

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports ICSharpCode.SharpZipLib
Imports Microsoft.Web.Administration

Namespace Contensive.Addons.AddonManager
    Public Class iisRecycleClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Using iisManager As ServerManager = New ServerManager()
                    Dim sites As SiteCollection = iisManager.Sites
                    For Each site As Site In sites
                        If (site.Name = CP.Site.Name) Then
                            iisManager.ApplicationPools(site.Applications("/").ApplicationPoolName).Recycle()
                        End If
                    Next
                End Using
                'Using myProcess As New Process
                '    myProcess.StartInfo.UseShellExecute = False
                '    myProcess.StartInfo.FileName = "c:\windows\system32\inetsrv\appcmd.exe recycle apppool """ & CP.Site.Name & """"
                '    myProcess.StartInfo.CreateNoWindow = True
                '    myProcess.Start()
                'End Using
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
            End Try
            Return returnHtml
        End Function
    End Class
End Namespace
