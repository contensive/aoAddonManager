#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]   $LocalDeployTarget  = '',
    [hashtable]$RemoteDeployTarget = $null
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\..\Contensive5\scripts\build-addon-collection.psm1') -Force

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path

Invoke-ContensiveBuild `
    -CollectionName    'Add-on Manager' `
    -CollectionPath    "$projectRoot\collections\Add-on Manager" `
    -SolutionPath      "$projectRoot\source\addonManager.sln" `
    -DotnetProjectPath "$projectRoot\source\addonManager51\addonManager51.csproj" `
    -BinPath           "$projectRoot\source\addonManager51\bin\Release\netstandard2.0" `
    -Configuration     'Release' `
    -DeploymentRoot    'C:\Deployments\aoAddonManager' `
    -CleanFolders      @(
                           "$projectRoot\source\addonManager51\bin"
                           "$projectRoot\source\addonManager51\obj"
                       ) `
    -UiPath            "$projectRoot\ui" `
    -UiAssetFolders    @('wwwFiles', 'cdnFiles', 'privateFiles', 'layoutFiles') `
    -LocalDeployTarget  $LocalDeployTarget `
    -RemoteDeployTarget $RemoteDeployTarget
