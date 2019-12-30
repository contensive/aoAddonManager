
Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.Addons.AddonManager51.Models
Imports Contensive.BaseClasses
Imports Contensive.Models.Db
Imports ICSharpCode.SharpZipLib

Namespace Contensive.Addons.AddonManager51
    Public Class ExportController
        '
        '====================================================================================================
        ''' <summary>
        ''' create the colleciton zip file and return the pathFilename in the Cdn
        ''' </summary>
        ''' <param name="cp"></param>
        ''' <param name="collectionId"></param>
        ''' <returns></returns>
        Public Shared Function createCollectionZip_returnCdnPathFilename(cp As CPBaseClass, collectionId As Integer) As String
            Dim cdnExportZip_Filename As String = ""
            Try
                Dim collection As AddonCollectionModel = DbBaseModel.create(Of AddonCollectionModel)(cp, collectionId)
                If (collection Is Nothing) Then
                    Call cp.UserError.Add("The collection you selected could not be found")
                Else
                End If
                Using CS As CPCSBaseClass = cp.CSNew()
                    CS.OpenRecord("Add-on Collections", collectionId)
                    If Not CS.OK() Then
                        Call cp.UserError.Add("The collection you selected could not be found")
                    Else
                        Dim collectionXml As String = "<?xml version=""1.0"" encoding=""windows-1252""?>"
                        Dim CollectionGuid As String = CS.GetText("ccGuid")
                        If CollectionGuid = "" Then
                            CollectionGuid = cp.Utils.CreateGuid()
                            Call CS.SetField("ccGuid", CollectionGuid)
                        End If
                        Dim onInstallAddonGuid As String = ""
                        If (CS.FieldOK("onInstallAddonId")) Then
                            Dim onInstallAddonId As Integer = CS.GetInteger("onInstallAddonId")
                            If (onInstallAddonId > 0) Then
                                Dim addon As AddonModel = AddonModel.create(Of AddonModel)(cp, onInstallAddonId)
                                If (addon IsNot Nothing) Then
                                    onInstallAddonGuid = addon.ccguid
                                End If
                            End If
                        End If
                        Dim CollectionName As String = CS.GetText("name")
                        collectionXml &= vbCrLf & "<Collection"
                        collectionXml &= " name=""" & CollectionName & """"
                        collectionXml &= " guid=""" & CollectionGuid & """"
                        collectionXml &= " system=""" & getYesNo(cp, CS.GetBoolean("system")) & """"
                        collectionXml &= " updatable=""" & getYesNo(cp, CS.GetBoolean("updatable")) & """"
                        collectionXml &= " blockNavigatorNode=""" & getYesNo(cp, CS.GetBoolean("blockNavigatorNode")) & """"
                        collectionXml &= " onInstallAddonGuid=""" & onInstallAddonGuid & """"
                        collectionXml &= ">"
                        cdnExportZip_Filename = encodeFilename(cp, CollectionName & ".zip")
                        Dim tempPathFileList As New List(Of String)
                        Dim tempExportPath As String = "CollectionExport" & Guid.NewGuid().ToString() & "\"
                        '
                        Dim resourceNodeList As String = ExportResourceListController.getResourceList(cp, CS.GetText("execFileList"), CollectionGuid, tempPathFileList, tempExportPath)
                        '
                        ' helpLink
                        '
                        If CS.FieldOK("HelpLink") Then
                            collectionXml &= vbCrLf & vbTab & "<HelpLink>" & System.Net.WebUtility.HtmlEncode(CS.GetText("HelpLink")) & "</HelpLink>"
                        End If
                        '
                        ' Help
                        '
                        collectionXml &= vbCrLf & vbTab & "<Help>" & System.Net.WebUtility.HtmlEncode(CS.GetText("Help")) & "</Help>"
                        '
                        ' Addons
                        '
                        Dim IncludeSharedStyleGuidList As String = ""
                        Dim IncludeModuleGuidList As String = ""
                        Using CS2 As CPCSBaseClass = cp.CSNew()
                            CS2.Open("Add-ons", "collectionid=" & collectionId, "name", True, "id")
                            Do While CS2.OK()
                                collectionXml &= getAddonNode(cp, CS2.GetInteger("id"), IncludeModuleGuidList, IncludeSharedStyleGuidList)
                                Call CS2.GoNext()
                            Loop
                        End Using
                        '
                        ' Data Records
                        '
                        Dim DataRecordList As String = CS.GetText("DataRecordList")
                        collectionXml &= DataRecordNodeListController.getNodeList(cp, DataRecordList, tempPathFileList, tempExportPath)
                        '
                        ' CDef
                        '

                        For Each content As Models.ContentModel In Models.ContentModel.createListFromCollection(cp, collectionId)
                            Dim reload As Boolean = False
                            If (String.IsNullOrEmpty(content.ccguid)) Then
                                content.ccguid = cp.Utils.CreateGuid()
                                content.save(cp)
                                reload = True
                            End If
                            Dim xmlTool As New xmlController(cp)
                            Dim Node As String = xmlTool.GetXMLContentDefinition3(content.name)
                            '
                            ' remove the <collection> top node
                            '
                            Dim Pos As Integer = InStr(1, Node, "<cdef", vbTextCompare)
                            If Pos > 0 Then
                                Node = Mid(Node, Pos)
                                Pos = InStr(1, Node, "</cdef>", vbTextCompare)
                                If Pos > 0 Then
                                    Node = Mid(Node, 1, Pos + 6)
                                    collectionXml &= vbCrLf & vbTab & Node
                                End If
                            End If
                        Next
                        '
                        ' Scripting Modules
                        '
                        'Call Main.testpoint("getCollection, 800")

                        If IncludeModuleGuidList <> "" Then
                            Dim Modules() As String = Split(IncludeModuleGuidList, vbCrLf)
                            For Ptr = 0 To UBound(Modules)
                                Dim ModuleGuid As String = Modules(Ptr)
                                If ModuleGuid <> "" Then
                                    Using CS2 As CPCSBaseClass = cp.CSNew()
                                        CS2.Open("Scripting Modules", "ccguid=" & cp.Db.EncodeSQLText(ModuleGuid))
                                        If CS2.OK() Then
                                            Dim Code As String = Trim(CS2.GetText("code"))
                                            Code = EncodeCData(cp, Code)
                                            collectionXml &= vbCrLf & vbTab & "<ScriptingModule Name=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("name")) & """ guid=""" & ModuleGuid & """>" & Code & "</ScriptingModule>"
                                        End If
                                        Call CS2.Close()
                                    End Using
                                End If
                            Next
                        End If
                        '
                        ' shared styles
                        '
                        Dim recordGuids() As String
                        Dim recordGuid As String
                        If (IncludeSharedStyleGuidList <> "") Then
                            recordGuids = Split(IncludeSharedStyleGuidList, vbCrLf)
                            For Ptr = 0 To UBound(recordGuids)
                                recordGuid = recordGuids(Ptr)
                                If recordGuid <> "" Then
                                    Using CS2 As CPCSBaseClass = cp.CSNew()
                                        CS2.Open("Shared Styles", "ccguid=" & cp.Db.EncodeSQLText(recordGuid))
                                        If CS2.OK() Then
                                            collectionXml &= vbCrLf & vbTab & "<SharedStyle" _
                                            & " Name=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("name")) & """" _
                                            & " guid=""" & recordGuid & """" _
                                            & " alwaysInclude=""" & CS2.GetBoolean("alwaysInclude") & """" _
                                            & " prefix=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("prefix")) & """" _
                                            & " suffix=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("suffix")) & """" _
                                            & " sortOrder=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("sortOrder")) & """" _
                                            & ">" _
                                            & EncodeCData(cp, Trim(CS2.GetText("styleFilename"))) _
                                            & "</SharedStyle>"
                                        End If
                                        Call CS2.Close()
                                    End Using
                                End If
                            Next
                        End If
                        '
                        ' Import Collections
                        '
                        If True Then
                            Dim Node As String = ""
                            Using CS3 As CPCSBaseClass = cp.CSNew()
                                If CS3.Open("Add-on Collection Parent Rules", "parentid=" & collectionId) Then
                                    Do
                                        Using CS2 As CPCSBaseClass = cp.CSNew()
                                            CS2.OpenRecord("Add-on Collections", CS3.GetInteger("childid"))
                                            If CS2.OK() Then
                                                Dim Guid As String = CS2.GetText("ccGuid")
                                                If Guid = "" Then
                                                    Guid = cp.Utils.CreateGuid()
                                                    Call CS2.SetField("ccGuid", Guid)
                                                End If
                                                Node = Node & vbCrLf & vbTab & "<ImportCollection name=""" & System.Net.WebUtility.HtmlEncode(CS2.GetText("name")) & """>" & Guid & "</ImportCollection>"
                                            End If
                                            Call CS2.Close()
                                        End Using
                                        Call CS3.GoNext()
                                    Loop While CS3.OK()
                                End If
                                Call CS3.Close()
                            End Using
                            collectionXml &= Node
                        End If
                        '
                        ' wwwFileList
                        '
                        If (True) Then
                            Dim wwwFileList As String = CS.GetText("wwwFileList")
                            If wwwFileList <> "" Then
                                Dim Files() As String = Split(wwwFileList, vbCrLf)
                                For Ptr = 0 To UBound(Files)
                                    Dim PathFilename As String = Files(Ptr)
                                    If PathFilename <> "" Then
                                        PathFilename = Replace(PathFilename, "\", "/")
                                        Dim Path As String = ""
                                        Dim Filename As String = PathFilename
                                        Dim Pos As Integer = InStrRev(PathFilename, "/")
                                        If Pos > 0 Then
                                            Filename = Mid(PathFilename, Pos + 1)
                                            Path = Mid(PathFilename, 1, Pos - 1)
                                        End If
                                        If LCase(Filename) = "collection.hlp" Then
                                            '
                                            ' legacy file, remove it
                                            '
                                        Else
                                            PathFilename = Replace(PathFilename, "/", "\")
                                            If tempPathFileList.Contains(tempExportPath & Filename) Then
                                                Call cp.UserError.Add("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                                            Else
                                                cp.WwwFiles.Copy(PathFilename, tempExportPath & Filename, cp.TempFiles)
                                                tempPathFileList.Add(tempExportPath & Filename)
                                                collectionXml &= vbCrLf & vbTab & "<Resource name=""" & System.Net.WebUtility.HtmlEncode(Filename) & """ type=""www"" path=""" & System.Net.WebUtility.HtmlEncode(Path) & """ />"
                                            End If
                                        End If
                                    End If
                                Next
                            End If

                        End If
                        '
                        ' ContentFileList
                        '
                        If True Then
                            Dim ContentFileList As String = CS.GetText("ContentFileList")
                            If ContentFileList <> "" Then
                                Dim Files() As String = Split(ContentFileList, vbCrLf)
                                For Ptr = 0 To UBound(Files)
                                    Dim PathFilename As String = Files(Ptr)
                                    If PathFilename <> "" Then
                                        PathFilename = Replace(PathFilename, "\", "/")
                                        Dim Path As String = ""
                                        Dim Filename As String = PathFilename
                                        Dim Pos As Integer = InStrRev(PathFilename, "/")
                                        If Pos > 0 Then
                                            Filename = Mid(PathFilename, Pos + 1)
                                            Path = Mid(PathFilename, 1, Pos - 1)
                                        End If
                                        If tempPathFileList.Contains(tempExportPath & Filename) Then
                                            Call cp.UserError.Add("There was an error exporting this collection because there were multiple files with the same filename [" & Filename & "]")
                                        Else
                                            cp.CdnFiles.Copy(PathFilename, tempExportPath + Filename, cp.TempFiles)
                                            tempPathFileList.Add(tempExportPath & Filename)
                                            collectionXml &= vbCrLf & vbTab & "<Resource name=""" & System.Net.WebUtility.HtmlEncode(Filename) & """ type=""content"" path=""" & System.Net.WebUtility.HtmlEncode(Path) & """ />"
                                        End If
                                    End If
                                Next
                            End If
                        End If
                        '
                        ' ExecFileListNode
                        '
                        collectionXml &= resourceNodeList
                        '
                        ' Other XML
                        '
                        Dim OtherXML As String
                        OtherXML = CS.GetText("otherxml")
                        If Trim(OtherXML) <> "" Then
                            collectionXml &= vbCrLf & OtherXML
                        End If
                        collectionXml &= vbCrLf & "</Collection>"
                        Call CS.Close()
                        Dim tempExportXml_Filename As String = encodeFilename(cp, CollectionName & ".xml")
                        '
                        ' Save the installation file and add it to the archive
                        '
                        Call cp.TempFiles.Save(tempExportPath & tempExportXml_Filename, collectionXml)
                        If Not tempPathFileList.Contains(tempExportPath & tempExportXml_Filename) Then
                            tempPathFileList.Add(tempExportPath & tempExportXml_Filename)
                        End If
                        Dim tempExportZip_Filename As String = encodeFilename(cp, CollectionName & ".zip")
                        '
                        ' -- zip up the folder to make the collection zip file in temp filesystem
                        Call zipTempCdnFile(cp, tempExportPath & tempExportZip_Filename, tempPathFileList)
                        '
                        ' -- copy the collection zip file to the cdn filesystem as the download link
                        cp.TempFiles.Copy(tempExportPath & tempExportZip_Filename, cdnExportZip_Filename, cp.CdnFiles)
                        '
                        ' -- delete the temp folder
                        cp.TempFiles.DeleteFolder(tempExportPath)
                    End If
                End Using
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
            Return cdnExportZip_Filename
        End Function
        '
        '====================================================================================================

        Public Shared Function getAddonNode(cp As CPBaseClass, addonid As Integer, ByRef Return_IncludeModuleGuidList As String, ByRef Return_IncludeSharedStyleGuidList As String) As String
            Dim result As String = ""
            Try
                Using CS As CPCSBaseClass = cp.CSNew()
                    If CS.OpenRecord("Add-ons", addonid) Then
                        Dim addonName As String = CS.GetText("name")
                        Dim processRunOnce As Boolean = CS.GetBoolean("ProcessRunOnce")
                        If ((LCase(addonName) = "oninstall") Or (LCase(addonName) = "_oninstall")) Then
                            processRunOnce = True
                        End If
                        '
                        ' ActiveX DLL node is being deprecated. This should be in the collection resource section
                        result &= getNodeText(cp, "Copy", CS.GetText("Copy"))
                        result &= getNodeText(cp, "CopyText", CS.GetText("CopyText"))
                        '
                        ' DLL
                        result &= getNodeText(cp, "ActiveXProgramID", CS.GetText("objectprogramid"), True)
                        result &= getNodeText(cp, "DotNetClass", CS.GetText("DotNetClass"))
                        '
                        ' Features
                        result &= getNodeText(cp, "ArgumentList", CS.GetText("ArgumentList"))
                        result &= getNodeBoolean(cp, "AsAjax", CS.GetBoolean("AsAjax"))
                        result &= getNodeBoolean(cp, "Filter", CS.GetBoolean("Filter"))
                        result &= getNodeText(cp, "Help", CS.GetText("Help"))
                        result &= getNodeText(cp, "HelpLink", CS.GetText("HelpLink"))
                        result &= vbCrLf & vbTab & "<Icon Link=""" & CS.GetText("iconfilename") & """ width=""" & CS.GetInteger("iconWidth") & """ height=""" & CS.GetInteger("iconHeight") & """ sprites=""" & CS.GetInteger("iconSprites") & """ />"
                        result &= getNodeBoolean(cp, "InIframe", CS.GetBoolean("InFrame"))
                        result &= If(CS.FieldOK("BlockEditTools"), getNodeBoolean(cp, "BlockEditTools", CS.GetBoolean("BlockEditTools")), "")
                        result &= If(CS.FieldOK("aliasList"), getNodeText(cp, "AliasList", CS.GetText("aliasList")), "")
                        '
                        ' -- Form XML
                        result &= getNodeText(cp, "FormXML", CS.GetText("FormXML"))
                        '
                        ' -- addon dependencies
                        Using CS2 As CPCSBaseClass = cp.CSNew()
                            CS2.Open("Add-on Include Rules", "addonid=" & addonid)
                            Do While CS2.OK()
                                Dim IncludedAddonID As Integer = CS2.GetInteger("IncludedAddonID")
                                Using CS3 As CPCSBaseClass = cp.CSNew()
                                    CS3.Open("Add-ons", "ID=" & IncludedAddonID)
                                    If CS3.OK() Then
                                        Dim Guid As String = CS3.GetText("ccGuid")
                                        If Guid = "" Then
                                            Guid = cp.Utils.CreateGuid()
                                            Call CS3.SetField("ccGuid", Guid)
                                        End If
                                        result &= vbCrLf & vbTab & "<IncludeAddon name=""" & System.Net.WebUtility.HtmlEncode(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                                    End If
                                    Call CS3.Close()
                                End Using
                                Call CS2.GoNext()
                            Loop
                            Call CS2.Close()
                        End Using
                        '
                        ' -- is inline/block
                        result &= getNodeBoolean(cp, "IsInline", CS.GetBoolean("IsInline"))
                        '
                        ' -- javascript (xmlnode may not match Db filename)
                        result &= getNodeText(cp, "JavascriptInHead", CS.GetText("JSFilename"))
                        result &= getNodeBoolean(cp, "javascriptForceHead", CS.GetBoolean("javascriptForceHead"))
                        result &= getNodeText(cp, "JSHeadScriptSrc", CS.GetText("JSHeadScriptSrc"))
                        '
                        ' -- javascript deprecated
                        result &= getNodeText(cp, "JSBodyScriptSrc", CS.GetText("JSBodyScriptSrc"), True)
                        result &= getNodeText(cp, "JavascriptBodyEnd", CS.GetText("JavascriptBodyEnd"), True)
                        result &= getNodeText(cp, "JavascriptOnLoad", CS.GetText("JavascriptOnLoad"), True)
                        '
                        ' -- Placements
                        result &= getNodeBoolean(cp, "Content", CS.GetBoolean("Content"))
                        result &= getNodeBoolean(cp, "Template", CS.GetBoolean("Template"))
                        result &= getNodeBoolean(cp, "Email", CS.GetBoolean("Email"))
                        result &= getNodeBoolean(cp, "Admin", CS.GetBoolean("Admin"))
                        result &= getNodeBoolean(cp, "OnPageEndEvent", CS.GetBoolean("OnPageEndEvent"))
                        result &= getNodeBoolean(cp, "OnPageStartEvent", CS.GetBoolean("OnPageStartEvent"))
                        result &= getNodeBoolean(cp, "OnBodyStart", CS.GetBoolean("OnBodyStart"))
                        result &= getNodeBoolean(cp, "OnBodyEnd", CS.GetBoolean("OnBodyEnd"))
                        result &= getNodeBoolean(cp, "RemoteMethod", CS.GetBoolean("RemoteMethod"))
                        result &= If(CS.FieldOK("Diagnostic"), getNodeBoolean(cp, "Diagnostic", CS.GetBoolean("Diagnostic")), "")
                        '
                        ' -- Process
                        result &= getNodeBoolean(cp, "ProcessRunOnce", processRunOnce)
                        result &= GetNodeInteger(cp, "ProcessInterval", CS.GetInteger("ProcessInterval"))
                        '
                        ' Meta
                        '
                        result &= getNodeText(cp, "MetaDescription", CS.GetText("MetaDescription"))
                        result &= getNodeText(cp, "OtherHeadTags", CS.GetText("OtherHeadTags"))
                        result &= getNodeText(cp, "PageTitle", CS.GetText("PageTitle"))
                        result &= getNodeText(cp, "RemoteAssetLink", CS.GetText("RemoteAssetLink"))
                        '
                        ' Styles
                        Dim Styles As String = ""
                        If Not CS.GetBoolean("BlockDefaultStyles") Then
                            Styles = Trim(CS.GetText("StylesFilename"))
                        End If
                        Dim StylesTest As String = Trim(CS.GetText("CustomStylesFilename"))
                        If StylesTest <> "" Then
                            If Styles <> "" Then
                                Styles = Styles & vbCrLf & StylesTest
                            Else
                                Styles = StylesTest
                            End If
                        End If
                        result &= getNodeText(cp, "Styles", Styles)
                        result &= getNodeText(cp, "styleslinkhref", CS.GetText("styleslinkhref"))
                        '
                        '
                        ' Scripting
                        '
                        Dim NodeInnerText As String = Trim(CS.GetText("ScriptingCode"))
                        If NodeInnerText <> "" Then
                            NodeInnerText = vbCrLf & vbTab & vbTab & "<Code>" & EncodeCData(cp, NodeInnerText) & "</Code>"
                        End If
                        Using CS2 As CPCSBaseClass = cp.CSNew()
                            CS2.Open("Add-on Scripting Module Rules", "addonid=" & addonid)
                            Do While CS2.OK()
                                Dim ScriptingModuleID As Integer = CS2.GetInteger("ScriptingModuleID")
                                Using CS3 As CPCSBaseClass = cp.CSNew()
                                    CS3.Open("Scripting Modules", "ID=" & ScriptingModuleID)
                                    If CS3.OK() Then
                                        Dim Guid As String = CS3.GetText("ccGuid")
                                        If Guid = "" Then
                                            Guid = cp.Utils.CreateGuid()
                                            Call CS3.SetField("ccGuid", Guid)
                                        End If
                                        Return_IncludeModuleGuidList = Return_IncludeModuleGuidList & vbCrLf & Guid
                                        NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<IncludeModule name=""" & System.Net.WebUtility.HtmlEncode(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                                    End If
                                    Call CS3.Close()
                                End Using
                                Call CS2.GoNext()
                            Loop
                            Call CS2.Close()
                        End Using
                        If NodeInnerText = "" Then
                            result &= vbCrLf & vbTab & "<Scripting Language=""" & CS.GetText("ScriptingLanguageID") & """ EntryPoint=""" & CS.GetText("ScriptingEntryPoint") & """ Timeout=""" & CS.GetText("ScriptingTimeout") & """/>"
                        Else
                            result &= vbCrLf & vbTab & "<Scripting Language=""" & CS.GetText("ScriptingLanguageID") & """ EntryPoint=""" & CS.GetText("ScriptingEntryPoint") & """ Timeout=""" & CS.GetText("ScriptingTimeout") & """>" & NodeInnerText & vbCrLf & vbTab & "</Scripting>"
                        End If
                        '
                        ' Shared Styles
                        '
                        Using CS2 As CPCSBaseClass = cp.CSNew()
                            CS2.Open("Shared Styles Add-on Rules", "addonid=" & addonid)
                            Do While CS2.OK()
                                Dim styleId As Integer = CS2.GetInteger("styleId")
                                Using CS3 As CPCSBaseClass = cp.CSNew()
                                    CS3.Open("shared styles", "ID=" & styleId)
                                    If CS3.OK() Then
                                        Dim Guid As String = CS3.GetText("ccGuid")
                                        If Guid = "" Then
                                            Guid = cp.Utils.CreateGuid()
                                            Call CS3.SetField("ccGuid", Guid)
                                        End If
                                        Return_IncludeSharedStyleGuidList = Return_IncludeSharedStyleGuidList & vbCrLf & Guid
                                        result &= vbCrLf & vbTab & "<IncludeSharedStyle name=""" & System.Net.WebUtility.HtmlEncode(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                                    End If
                                    Call CS3.Close()
                                End Using
                                Call CS2.GoNext()
                            Loop
                            Call CS2.Close()
                        End Using
                        '
                        ' Process Triggers
                        '
                        NodeInnerText = ""
                        Using CS2 As CPCSBaseClass = cp.CSNew()
                            CS2.Open("Add-on Content Trigger Rules", "addonid=" & addonid)
                            Do While CS2.OK()
                                Dim TriggerContentID As Integer = CS2.GetInteger("ContentID")
                                Using CS3 As CPCSBaseClass = cp.CSNew()
                                    CS3.Open("content", "ID=" & TriggerContentID)
                                    If CS3.OK() Then
                                        Dim Guid As String = CS3.GetText("ccGuid")
                                        If Guid = "" Then
                                            Guid = cp.Utils.CreateGuid()
                                            Call CS3.SetField("ccGuid", Guid)
                                        End If
                                        NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<ContentChange name=""" & System.Net.WebUtility.HtmlEncode(CS3.GetText("name")) & """ guid=""" & Guid & """/>"
                                    End If
                                    Call CS3.Close()
                                End Using
                                Call CS2.GoNext()
                            Loop
                            Call CS2.Close()
                        End Using
                        If NodeInnerText <> "" Then
                            result &= vbCrLf & vbTab & "<ProcessTriggers>" & NodeInnerText & vbCrLf & vbTab & "</ProcessTriggers>"
                        End If
                        '
                        ' Editors
                        '
                        If cp.Content.IsField("Add-on Content Field Type Rules", "id") Then
                            NodeInnerText = ""
                            Using CS2 As CPCSBaseClass = cp.CSNew()
                                CS2.Open("Add-on Content Field Type Rules", "addonid=" & addonid)
                                Do While CS2.OK()
                                    Dim fieldTypeID As Integer = CS2.GetInteger("contentFieldTypeID")
                                    Dim fieldType As String = cp.Content.GetRecordName("Content Field Types", fieldTypeID)
                                    If fieldType <> "" Then
                                        NodeInnerText = NodeInnerText & vbCrLf & vbTab & vbTab & "<type>" & fieldType & "</type>"
                                    End If
                                    Call CS2.GoNext()
                                Loop
                                Call CS2.Close()
                            End Using
                            If NodeInnerText <> "" Then
                                result &= vbCrLf & vbTab & "<Editors>" & NodeInnerText & vbCrLf & vbTab & "</Editors>"
                            End If
                        End If
                        '
                        Dim addonGuid As String = CS.GetText("ccGuid")
                        If (String.IsNullOrWhiteSpace(addonGuid)) Then
                            addonGuid = cp.Utils.CreateGuid()
                            Call CS.SetField("ccGuid", addonGuid)
                        End If
                        Dim NavType As String = CS.GetText("NavTypeID")
                        If (NavType = "") Then
                            NavType = "Add-on"
                        End If
                        result = "" _
                        & vbCrLf & vbTab & "<Addon name=""" & System.Net.WebUtility.HtmlEncode(addonName) & """ guid=""" & addonGuid & """ type=""" & NavType & """>" _
                        & tabIndent(cp, result) _
                        & vbCrLf & vbTab & "</Addon>"
                    End If
                    Call CS.Close()

                End Using
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetAddonNode")
            End Try
            Return result
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple text node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="NodeContent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getNodeText(cp As CPBaseClass, NodeName As String, NodeContent As String, Optional deprecated As Boolean = False) As String
            getNodeText = ""
            Try
                Dim prefix As String = ""
                If (deprecated) Then
                    prefix = "<!-- deprecated -->"
                End If
                getNodeText = ""
                If NodeContent = "" Then
                    getNodeText = getNodeText & vbCrLf & vbTab & prefix & "<" & NodeName & "></" & NodeName & ">"
                Else
                    getNodeText = getNodeText & vbCrLf & vbTab & prefix & "<" & NodeName & ">" & EncodeCData(cp, NodeContent) & "</" & NodeName & ">"
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "getNodeText")
            End Try
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple boolean node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="NodeContent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getNodeBoolean(cp As CPBaseClass, NodeName As String, NodeContent As Boolean) As String
            getNodeBoolean = ""
            Try
                getNodeBoolean = vbCrLf & vbTab & "<" & NodeName & ">" & getYesNo(cp, NodeContent) & "</" & NodeName & ">"
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetNodeBoolean")
            End Try
        End Function
        '
        '====================================================================================================
        ''' <summary>
        ''' create a simple integer node with a name and content
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <param name="NodeContent"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetNodeInteger(cp As CPBaseClass, NodeName As String, NodeContent As Integer) As String
            GetNodeInteger = ""
            Try
                GetNodeInteger = vbCrLf & vbTab & "<" & NodeName & ">" & CStr(NodeContent) & "</" & NodeName & ">"
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetNodeInteger")
            End Try
        End Function
        '
        '====================================================================================================
        Public Shared Function replaceMany(cp As CPBaseClass, Source As String, ArrayOfSource() As String, ArrayOfReplacement() As String) As String
            replaceMany = ""
            Try
                Dim Count As Integer = UBound(ArrayOfSource) + 1
                replaceMany = Source
                For Pointer = 0 To Count - 1
                    replaceMany = Replace(replaceMany, ArrayOfSource(Pointer), ArrayOfReplacement(Pointer))
                Next
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "replaceMany")
            End Try
        End Function
        '
        '====================================================================================================
        Public Shared Function encodeFilename(cp As CPBaseClass, Filename As String) As String
            encodeFilename = ""
            Try
                Dim Source() As String
                Dim Replacement() As String
                '
                Source = {"""", "*", "/", ":", "<", ">", "?", "\", "|"}
                Replacement = {"_", "_", "_", "_", "_", "_", "_", "_", "_"}
                '
                encodeFilename = replaceMany(cp, Filename, Source, Replacement)
                If Len(encodeFilename) > 254 Then
                    encodeFilename = Left(encodeFilename, 254)
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "encodeFilename")
            End Try
        End Function
        '
        '====================================================================================================
        '
        Public Shared Sub GetLocalCollectionArgs(cp As CPBaseClass, CollectionGuid As String, ByRef Return_CollectionPath As String, ByRef Return_LastChagnedate As Date)
            Try
                Const CollectionListRootNode = "collectionlist"
                '
                Dim LocalPath As String
                Dim LocalGuid As String = ""
                Dim Doc As New Xml.XmlDocument
                Dim CollectionNode As Xml.XmlNode
                Dim LocalListNode As Xml.XmlNode
                Dim CollectionFound As Boolean
                Dim CollectionPath As String = ""
                Dim LastChangeDate As Date
                Dim MatchFound As Boolean
                Dim LocalName As String
                '
                MatchFound = False
                Return_CollectionPath = ""
                Return_LastChagnedate = Date.MinValue
                Call Doc.LoadXml(cp.PrivateFiles.Read("addons\Collections.xml"))
                If True Then
                    If LCase(Doc.DocumentElement.Name) <> LCase(CollectionListRootNode) Then
                        'Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, Hint=[" & Hint & "], The Collections.xml file has an invalid root node, [" & Doc.documentElement.name & "] was received and [" & CollectionListRootNode & "] was expected.")
                    Else
                        With Doc.DocumentElement
                            If LCase(.Name) <> "collectionlist" Then
                                'Call AppendClassLogFile("Server", "", "GetLocalCollectionArgs, basename was not collectionlist, [" & .name & "].")
                            Else
                                CollectionFound = False
                                'hint = hint & ",checking nodes [" & .childNodes.length & "]"
                                For Each LocalListNode In .ChildNodes
                                    LocalName = "no name found"
                                    LocalPath = ""
                                    Select Case LCase(LocalListNode.Name)
                                        Case "collection"
                                            LocalGuid = ""
                                            For Each CollectionNode In LocalListNode.ChildNodes
                                                Select Case LCase(CollectionNode.Name)
                                                    Case "name"
                                                        '
                                                        LocalName = LCase(CollectionNode.InnerText)
                                                    Case "guid"
                                                        '
                                                        LocalGuid = LCase(CollectionNode.InnerText)
                                                    Case "path"
                                                        '
                                                        CollectionPath = LCase(CollectionNode.InnerText)
                                                    Case "lastchangedate"
                                                        LastChangeDate = cp.Utils.EncodeDate(CollectionNode.InnerText)
                                                End Select
                                            Next
                                    End Select
                                    'hint = hint & ",checking node [" & LocalName & "]"
                                    If LCase(CollectionGuid) = LocalGuid Then
                                        Return_CollectionPath = CollectionPath
                                        Return_LastChagnedate = LastChangeDate
                                        'Call AppendClassLogFile("Server", "GetCollectionConfigArg", "GetLocalCollectionArgs, match found, CollectionName=" & LocalName & ", CollectionPath=" & CollectionPath & ", LastChangeDate=" & LastChangeDate)
                                        MatchFound = True
                                        Exit For
                                    End If
                                Next
                            End If
                        End With
                    End If
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "GetLocalCollectionArgs")
            End Try
        End Sub
        '
        '====================================================================================================
        '
        Public Shared Function EncodeCData(cp As CPBaseClass, Source As String) As String
            EncodeCData = ""
            Try
                EncodeCData = Source
                If EncodeCData <> "" Then
                    EncodeCData = "<![CDATA[" & Replace(EncodeCData, "]]>", "]]]]><![CDATA[>") & "]]>"
                End If
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "EncodeCData")
            End Try
        End Function
        '
        '====================================================================================================
        Public Shared Function getYesNo(cp As CPBaseClass, Key As Boolean) As String
            Return If(Key, "Yes", "No")
        End Function
        '
        '=======================================================================================
        ''' <summary>
        ''' zip
        ''' </summary>
        ''' <param name="PathFilename"></param>
        ''' <remarks></remarks>
        Public Shared Sub UnzipFile(cp As CPBaseClass, ByVal PathFilename As String)
            Try
                '
                Dim fastZip As ICSharpCode.SharpZipLib.Zip.FastZip = New ICSharpCode.SharpZipLib.Zip.FastZip()
                Dim fileFilter As String = Nothing

                fastZip.ExtractZip(PathFilename, getPath(cp, PathFilename), fileFilter)                '
            Catch ex As Exception
                cp.Site.ErrorReport(ex, "UnzipFile")
            End Try
        End Sub        '
        '
        '=======================================================================================
        ''' <summary>
        ''' unzip
        ''' </summary>
        ''' <param name="zipTempPathFilename"></param>
        ''' <param name="addTempPathFilename"></param>
        ''' <remarks></remarks>
        Public Shared Sub zipTempCdnFile(cp As CPBaseClass, zipTempPathFilename As String, ByVal addTempPathFilename As List(Of String))
            Try
                Dim z As Zip.ZipFile
                If cp.TempFiles.FileExists(zipTempPathFilename) Then
                    '
                    ' update existing zip with list of files
                    z = New Zip.ZipFile(cp.TempFiles.PhysicalFilePath & zipTempPathFilename)
                Else
                    '
                    ' create new zip
                    z = Zip.ZipFile.Create(cp.TempFiles.PhysicalFilePath & zipTempPathFilename)
                End If
                z.BeginUpdate()
                For Each pathFilename In addTempPathFilename
                    z.Add(cp.TempFiles.PhysicalFilePath & pathFilename, System.IO.Path.GetFileName(pathFilename))
                Next
                z.CommitUpdate()
                z.Close()
            Catch ex As Exception
                cp.Site.ErrorReport(ex)
            End Try
        End Sub
        '
        '=======================================================================================
        '
        Public Shared Function getPath(cp As CPBaseClass, ByVal pathFilename As String) As String
            Dim Position As Integer = InStrRev(pathFilename, "\")
            If Position <> 0 Then
                Return Mid(pathFilename, 1, Position)
            End If
            Return String.Empty
        End Function
        '
        '=======================================================================================
        '
        Public Shared Function getFilename(cp As CPBaseClass, ByVal PathFilename As String) As String
            Dim pos As Integer = InStrRev(PathFilename, "/")
            If pos <> 0 Then
                Return Mid(PathFilename, pos + 1)
            End If
            Return PathFilename
        End Function
        '
        '=======================================================================================
        '
        '   Indent every line by 1 tab
        '
        Public Shared Function tabIndent(cp As CPBaseClass, Source As String) As String
            Dim posStart As Integer = InStr(1, Source, "<![CDATA[", CompareMethod.Text)
            If posStart = 0 Then
                '
                ' no cdata
                posStart = InStr(1, Source, "<textarea", CompareMethod.Text)
                If posStart = 0 Then
                    '
                    ' no textarea
                    '
                    tabIndent = Replace(Source, vbCrLf & vbTab, vbCrLf & vbTab & vbTab)
                Else
                    '
                    ' text area found, isolate it and indent before and after
                    '
                    Dim posEnd As Integer = InStr(posStart, Source, "</textarea>", CompareMethod.Text)
                    Dim pre As String = Mid(Source, 1, posStart - 1)
                    Dim post As String = ""
                    Dim target As String
                    If posEnd = 0 Then
                        target = Mid(Source, posStart)
                    Else
                        target = Mid(Source, posStart, posEnd - posStart + Len("</textarea>"))
                        post = Mid(Source, posEnd + Len("</textarea>"))
                    End If
                    tabIndent = tabIndent(cp, pre) & target & tabIndent(cp, post)
                End If
            Else
                '
                ' cdata found, isolate it and indent before and after
                '
                Dim posEnd As Integer = InStr(posStart, Source, "]]>", CompareMethod.Text)
                Dim pre As String = Mid(Source, 1, posStart - 1)
                Dim post As String = ""
                Dim target As String
                If posEnd = 0 Then
                    target = Mid(Source, posStart)
                Else
                    target = Mid(Source, posStart, posEnd - posStart + Len("]]>"))
                    post = Mid(Source, posEnd + 3)
                End If
                tabIndent = tabIndent(cp, pre) & target & tabIndent(cp, post)
            End If
        End Function
    End Class
End Namespace