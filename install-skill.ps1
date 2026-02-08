# Instala el skill vbaExcel en el directorio de Copilot.

$ErrorActionPreference = "Stop"

function Test-ExcelInstall {
    $versions = @("16.0", "15.0", "14.0")
    foreach ($ver in $versions) {
        $path = "HKLM:\SOFTWARE\Microsoft\Office\$ver\Excel\InstallRoot"
        if (Test-Path $path) { return $true }
    }
    return $false
}

function Get-AccessVBOMStatus {
    $versions = @("16.0", "15.0", "14.0")
    foreach ($ver in $versions) {
        $path = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
        if (Test-Path $path) {
            $value = (Get-ItemProperty -Path $path -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
            if ($value -eq 1) { return $true }
        }
    }
    return $false
}

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$bundlePath = Join-Path $repoRoot "skill-bundle"
$targetRoot = Join-Path $env:USERPROFILE ".copilot\skills"
$targetPath = Join-Path $targetRoot "vbaExcel"

if (-not (Test-Path $bundlePath)) {
    Write-Host "skill-bundle not found: $bundlePath" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $targetRoot)) {
    New-Item -ItemType Directory -Path $targetRoot | Out-Null
}

if (-not (Test-ExcelInstall)) {
    Write-Host "Excel does not appear to be installed." -ForegroundColor Yellow
    Write-Host "Install Excel before using this skill." -ForegroundColor Yellow
}

if (-not (Get-AccessVBOMStatus)) {
    Write-Host "Access to VBA project is not enabled." -ForegroundColor Yellow
    $regPath = Join-Path $bundlePath "scripts\enable_vba_access.reg"
    if (Test-Path $regPath) {
        $answer = Read-Host "Enable AccessVBOM now? (y/n)"
        if ($answer.ToLower() -eq "y") {
            try {
                reg import $regPath | Out-Null
                Write-Host "AccessVBOM enabled." -ForegroundColor Green
            } catch {
                Write-Host "Could not import registry file. Run manually:" -ForegroundColor Yellow
                Write-Host "reg import `"$regPath`"" -ForegroundColor Yellow
            }
        }
    }
}

if (Test-Path $targetPath) {
    $answer = Read-Host "Target exists. Overwrite? (y/n)"
    if ($answer.ToLower() -ne "y") {
        Write-Host "Install cancelled." -ForegroundColor Yellow
        exit 0
    }
    Remove-Item -Recurse -Force $targetPath
}

Copy-Item -Recurse -Force $bundlePath $targetPath

Write-Host "Skill installed to: $targetPath" -ForegroundColor Green
Write-Host "Restart VS Code to load the skill." -ForegroundColor Green
