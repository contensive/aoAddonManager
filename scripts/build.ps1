#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Debug','Release')]
    [string]   $Configuration      = 'Release',
    [string]   $LocalDeployTarget  = '',
    [hashtable]$RemoteDeployTarget = $null
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot '..\..\Contensive5\scripts\build-addon-collection.psm1') -Force

$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path

# Display build info and pause for confirmation
$version = Get-ContensiveVersion -DeploymentRoot 'C:\Deployments\aoAddonManager'
Write-Host ""
Write-Host "========================================"
Write-Host "  Add-on Manager Build"
Write-Host "  Version:       $version"
Write-Host "  Configuration: $Configuration"
Write-Host "========================================"
Write-Host ""
Write-Host "Press any key to continue or Ctrl+C to cancel..."
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
Write-Host ""

Invoke-ContensiveBuild `
    -CollectionName    'Add-on Manager' `
    -CollectionPath    "$projectRoot\collections\Add-on Manager" `
    -SolutionPath      "$projectRoot\source\addonManager.sln" `
    -DotnetProjectPath "$projectRoot\source\addonManager51\addonManager51.csproj" `
    -BinPath           "$projectRoot\source\addonManager51\bin\$Configuration\netstandard2.0" `
    -Configuration     $Configuration `
    -DeploymentRoot    'C:\Deployments\aoAddonManager' `
    -CleanFolders      @(
                           "$projectRoot\source\addonManager51\bin"
                           "$projectRoot\source\addonManager51\obj"
                       ) `
    -UiPath            "$projectRoot\ui" `
    -UiAssetFolders    @('wwwFiles', 'cdnFiles', 'privateFiles', 'layoutFiles') `
    -LocalDeployTarget  $LocalDeployTarget `
    -RemoteDeployTarget $RemoteDeployTarget
