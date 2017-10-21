VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FromFolderClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Dim Csv As Object
    Dim main As Object
    '
    Set Csv = CsvObject
    Set main = MainObject
    '
    Dim InstalledCollectionID As Long
    Dim InstalledCollectionGuid As String
    Dim ErrorMessage As String
    Dim IISResetRequired As Boolean
    Dim AddonInstall As Object
    Dim InstallPath As String
    Dim AdminUI As Object
    Dim UpgradeOK As Boolean
    '
    InstallPath = main.GetStreamText("installPath")
    Set AddonInstall = CreateObject("ccUpgrade.AddonInstallClass")
    Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
    '
    UpgradeOK = AddonInstall.InstallCollectionFilesFromFolder3(InstallPath, IISResetRequired, main.ApplicationName, ErrorMessage, InstalledCollectionGuid)
    If Not UpgradeOK Then
        If ErrorMessage = "" Then
            main.AddUserError ("The Add-on Collection did not install correctly, but no detailed error message was given.")
        Else
            main.AddUserError ("The Add-on Collection did not install correctly, " & ErrorMessage)
        End If
    End If
End Function


