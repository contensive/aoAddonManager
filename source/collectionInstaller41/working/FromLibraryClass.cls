VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FromLibraryClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Public Function Execute(Csv As Object, main As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    '
    Dim AddonInstall As Object
    Dim collectionGuid As String
    Dim IISResetRequired As Boolean
    Dim RegisterList As String
    Dim ErrorMessage As String
    '
    collectionGuid = GetAddonOptionStringValue("collectionGuid", OptionString)
    Set AddonInstall = CreateObject("ccUpgrade.AddonInstallClass")
    Execute = AddonInstall.UpgradeAllAppsFromLibCollection(collectionGuid, main.ApplicationName, IISResetRequired, RegisterList, ErrorMessage)
    '
    Exit Function
ErrorTrap:
    If main Is Nothing Then
        Call HandleError2("(main is nothing so app.name unknown)", "error trap", "ccWeb3", "AddonManClass", "", Err.Number, Err.Source, Err.Description, True, False, "unknown")
    Else
        Call HandleError2(main.ApplicationName, "error trap", "ccWeb3", "AddonManClass", "", Err.Number, Err.Source, Err.Description, True, False, main.ServerLink)
    End If
End Function
