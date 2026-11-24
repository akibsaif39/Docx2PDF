param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $false)]
    [string]$Notes = "Performance enhancements, conversion accuracy fixes, and stability improvements.",

    [Parameter(Mandatory = $false)]
    [string]$Token = $env:GITHUB_TOKEN
)

if ([string]::IsNullOrWhiteSpace($Token)) {
    $Token = $env:GH_TOKEN
}
if ([string]::IsNullOrWhiteSpace($Token)) {
    # Check for token in parent or local private config if present
    $tokenFile = Join-Path $PSScriptRoot ".token"
    if (Test-Path $tokenFile) {
        $Token = (Get-Content $tokenFile -Raw).Trim()
    }
}

$ErrorActionPreference = "Stop"

$cleanVer = $Version.Trim().TrimStart('v', 'V')
$tag = "v$cleanVer"
$owner = "akibsaif39"
$repo = "Docx2PDF"
$rootDir = (Get-Item $PSScriptRoot).Parent.FullName

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Docx2PDF Release Automation: $tag" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# 1. Update project version in Docx2PDFLive.UI.csproj and Product.wxs
Write-Host "`n[1/7] Updating version in Docx2PDFLive.UI.csproj & Product.wxs..." -ForegroundColor Yellow
$csprojPath = Join-Path $rootDir "Docx2PDFLive.UI\Docx2PDFLive.UI.csproj"
$csproj = Get-Content $csprojPath -Raw
$csproj = $csproj -replace '<Version>.*?</Version>', "<Version>$cleanVer</Version>"
$csproj = $csproj -replace '<AssemblyVersion>.*?</AssemblyVersion>', "<AssemblyVersion>$cleanVer.0</AssemblyVersion>"
$csproj = $csproj -replace '<FileVersion>.*?</FileVersion>', "<FileVersion>$cleanVer.0</FileVersion>"
Set-Content -Path $csprojPath -Value $csproj -NoNewline

$wxsPath = Join-Path $rootDir "Installer\Product.wxs"
$wxs = Get-Content $wxsPath -Raw
$wxs = $wxs -replace 'Version="[0-9\.]+"', "Version=""$cleanVer"""
Set-Content -Path $wxsPath -Value $wxs -NoNewline

# 2. Update local version.json
Write-Host "`n[2/7] Updating version.json..." -ForegroundColor Yellow
$versionJsonPath = Join-Path $rootDir "version.json"
$downloadUrl = "https://github.com/$owner/$repo/releases/download/$tag/Docx2PDF_Install.exe"
$versionJsonContent = @{
    latestVersion = $cleanVer
    mandatory = $true
    releaseNotes = $Notes
    downloadUrl = $downloadUrl
} | ConvertTo-Json -Depth 4
Set-Content -Path $versionJsonPath -Value $versionJsonContent

# 3. Publish standalone single-file Docx2PDF.exe
Write-Host "`n[3/7] Publishing standalone Docx2PDF.exe..." -ForegroundColor Yellow
$publishDir = Join-Path $rootDir "publish"
dotnet publish (Join-Path $rootDir "Docx2PDFLive.UI\Docx2PDFLive.UI.csproj") -c Release -o $publishDir

# Synchronize .dev\Docx2PDF.exe (if not currently locked by running instance)
$devExe = Join-Path $rootDir ".dev\Docx2PDF.exe"
try {
    Copy-Item (Join-Path $publishDir "Docx2PDF.exe") -Destination $devExe -Force -ErrorAction Stop
} catch {
    Write-Host "Note: .dev\Docx2PDF.exe is currently running. Skipping local dev file overwrite." -ForegroundColor Yellow
}

# 4. Rebuild native 64-bit WiX MSI
Write-Host "`n[4/7] Rebuilding native 64-bit MSI..." -ForegroundColor Yellow
$installerDir = Join-Path $rootDir "Installer"
Push-Location $installerDir
try {
    wix build -arch x64 Product.wxs -o Docx2PDF_Setup.msi
} finally {
    Pop-Location
}

# 5. Build unified Docx2PDF_Install.exe
Write-Host "`n[5/7] Building unified Docx2PDF_Install.exe..." -ForegroundColor Yellow
$setupPubDir = Join-Path $rootDir "scratch\setuppub"
dotnet publish (Join-Path $rootDir "Docx2PDFLive.Setup\Docx2PDFLive.Setup.csproj") -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o $setupPubDir
$rootInstaller = Join-Path $rootDir "Docx2PDF_Install.exe"
Copy-Item (Join-Path $setupPubDir "Docx2PDF_Install.exe") -Destination $rootInstaller -Force

# 6. Publish to GitHub Releases
Write-Host "`n[6/7] Creating/Updating GitHub Release $tag..." -ForegroundColor Yellow
$headers = @{
    "Authorization" = "Bearer $Token"
    "Accept" = "application/vnd.github.v3+json"
    "User-Agent" = "Docx2PDF-Manager"
}

# Check if release exists
$existingRelease = $null
try {
    $existingRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/releases/tags/$tag" -Headers $headers -Method Get
} catch { }

$releaseId = $null
if ($null -ne $existingRelease) {
    $releaseId = $existingRelease.id
    Write-Host "Found existing release $tag (ID: $releaseId). Cleaning up old assets..."
    foreach ($asset in $existingRelease.assets) {
        Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/releases/assets/$($asset.id)" -Headers $headers -Method Delete
    }
} else {
    Write-Host "Creating new GitHub release for $tag..."
    $releasePayload = @{
        tag_name = $tag
        name = "Docx2PDF $tag"
        body = "# Docx2PDF $tag`n`n$Notes`n`n### Download (Single Installer)`n👉 **[Download Docx2PDF_Install.exe]($downloadUrl)**"
        draft = $false
        prerelease = $false
    } | ConvertTo-Json
    $newRel = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/releases" -Headers $headers -Method Post -Body $releasePayload
    $releaseId = $newRel.id
}

Write-Host "Uploading $rootInstaller to release..."
$uploadUrl = "https://uploads.github.com/repos/$owner/$repo/releases/$releaseId/assets?name=Docx2PDF_Install.exe"
& curl.exe -X POST -H "Authorization: Bearer $Token" -H "Content-Type: application/octet-stream" --data-binary "@$rootInstaller" "$uploadUrl"

# 7. Commit and push source files to main branch
Write-Host "`n[7/7] Pushing updated source code and version.json to GitHub main..." -ForegroundColor Yellow
$pushScript = Join-Path $rootDir "scratch\push_to_github.ps1"
if (Test-Path $pushScript) {
    & powershell -ExecutionPolicy Bypass -File $pushScript -CommitMessage "Release Docx2PDF $tag`n`n$Notes"
}

Write-Host "`n=======================================================" -ForegroundColor Green
Write-Host " SUCCESS! Docx2PDF $tag is live!" -ForegroundColor Green
Write-Host " All running/installed apps will now auto-update." -ForegroundColor Green
Write-Host "=======================================================" -ForegroundColor Green
