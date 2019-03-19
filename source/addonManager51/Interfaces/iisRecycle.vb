

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports ICSharpCode.SharpZipLib

Namespace Contensive.Addons.AddonManager51
    Public Class iisRecycleClass
        Inherits AddonBaseClass
        '
        '====================================================================================================
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String = ""
            Try
                Using myProcess As New Process
                    myProcess.StartInfo.UseShellExecute = False
                    myProcess.StartInfo.FileName = "c:\windows\system32\inetsrv\appcmd recycle apppool """ & CP.Site.Name & """"
                    myProcess.StartInfo.CreateNoWindow = True
                    myProcess.Start()
                End Using
            Catch ex As Exception
                CP.Site.ErrorReport(ex)
            End Try
            Return returnHtml
        End Function
    End Class
End Namespace
