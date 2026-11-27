param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [Parameter(Mandatory = $false)]
    [string]$Notes = "Performance enhancements, conversion accuracy fixes, and stability improvements.",

    [Parameter(Mandatory = $false)]
    [string]$Token = $env:GITHUB_TOKEN
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($Token)) {
    $Token = $env:GH_TOKEN
}
if ([string]::IsNullOrWhiteSpace($Token)) {
    $tokenFile = Join-Path $PSScriptRoot ".token"
    if (Test-Path $tokenFile) {
        $Token = (Get-Content $tokenFile -Raw).Trim()
    }
}
if ([string]::IsNullOrWhiteSpace($Token)) {
    throw "GitHub token not found. Set `$env:GITHUB_TOKEN, pass -Token, or create scripts\.token."
}

$cleanVer = $Version.Trim().TrimStart('v', 'V')
if ($cleanVer -notmatch '^\d+\.\d+\.\d+$') {
    throw "Version '$Version' must follow Major.Minor.Build format (e.g. 1.0.1)."
}

$tag = "v$cleanVer"
$owner = "akibsaif39"
$repo = "Docx2PDF"
$rootDir = (Get-Item $PSScriptRoot).Parent.FullName
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

$headers = @{
    "Authorization" = "Bearer $Token"
    "Accept"        = "application/vnd.github.v3+json"
    "User-Agent"    = "Docx2PDF-Manager"
}

function Invoke-GitHubJson {
    param(
        [string]$Uri,
        [string]$Method = "Get",
        [object]$Body = $null
    )
    if ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 10
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        return Invoke-RestMethod -Uri $Uri -Headers $headers -Method $Method -ContentType "application/json; charset=utf-8" -Body $bytes
    }
    return Invoke-RestMethod -Uri $Uri -Headers $headers -Method $Method
}

function Get-GitBlobSha {
    param([byte[]]$Bytes)
    $header = [System.Text.Encoding]::ASCII.GetBytes("blob $($Bytes.Length)`0")
    $combined = New-Object byte[] ($header.Length + $Bytes.Length)
    [Buffer]::BlockCopy($header, 0, $combined, 0, $header.Length)
    [Buffer]::BlockCopy($Bytes, 0, $combined, $header.Length, $Bytes.Length)
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        return ([BitConverter]::ToString($sha1.ComputeHash($combined)) -replace '-', '').ToLowerInvariant()
    } finally {
        $sha1.Dispose()
    }
}

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
[System.IO.File]::WriteAllText($csprojPath, $csproj, $utf8NoBom)

$wxsPath = Join-Path $rootDir "Installer\Product.wxs"
$wxs = Get-Content $wxsPath -Raw
$wxs = $wxs -replace '(<Package\b[^>]*?\bVersion=")[0-9\.]+(")', "`${1}$cleanVer`${2}"
[System.IO.File]::WriteAllText($wxsPath, $wxs, $utf8NoBom)

# 2. Update local version.json
Write-Host "`n[2/7] Updating version.json..." -ForegroundColor Yellow
$versionJsonPath = Join-Path $rootDir "version.json"
$downloadUrl = "https://github.com/$owner/$repo/releases/download/$tag/Docx2PDF_Install.exe"
$versionJsonContent = @{
    latestVersion = $cleanVer
    mandatory     = $true
    releaseNotes  = $Notes
    downloadUrl   = $downloadUrl
} | ConvertTo-Json -Depth 4
[System.IO.File]::WriteAllText($versionJsonPath, ($versionJsonContent + "`n"), $utf8NoBom)

# 3. Publish standalone single-file Docx2PDF.exe
Write-Host "`n[3/7] Publishing standalone Docx2PDF.exe..." -ForegroundColor Yellow
$publishDir = Join-Path $rootDir "publish"
& dotnet publish (Join-Path $rootDir "Docx2PDFLive.UI\Docx2PDFLive.UI.csproj") -c Release -o $publishDir
if ($LASTEXITCODE -ne 0) { throw "dotnet publish Docx2PDFLive.UI failed with exit code $LASTEXITCODE." }

$devDir = Join-Path $rootDir ".dev"
if (Test-Path $devDir) {
    $devExe = Join-Path $devDir "Docx2PDF.exe"
    try {
        Copy-Item (Join-Path $publishDir "Docx2PDF.exe") -Destination $devExe -Force -ErrorAction Stop
    } catch {
        Write-Host "Note: .dev\Docx2PDF.exe is currently running. Skipping local dev file overwrite." -ForegroundColor Yellow
    }
}

# 4. Rebuild native 64-bit WiX MSI
Write-Host "`n[4/7] Rebuilding native 64-bit MSI..." -ForegroundColor Yellow
$installerDir = Join-Path $rootDir "Installer"
Push-Location $installerDir
try {
    & wix build -arch x64 Product.wxs -o Docx2PDF_Setup.msi
    if ($LASTEXITCODE -ne 0) { throw "wix build failed with exit code $LASTEXITCODE." }
} finally {
    Pop-Location
}

# 5. Build unified Docx2PDF_Install.exe
Write-Host "`n[5/7] Building unified Docx2PDF_Install.exe..." -ForegroundColor Yellow
$setupPubDir = Join-Path $rootDir "scratch\setuppub"
& dotnet publish (Join-Path $rootDir "Docx2PDFLive.Setup\Docx2PDFLive.Setup.csproj") -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o $setupPubDir
if ($LASTEXITCODE -ne 0) { throw "dotnet publish Docx2PDFLive.Setup failed with exit code $LASTEXITCODE." }

$rootInstaller = Join-Path $rootDir "Docx2PDF_Install.exe"
Copy-Item (Join-Path $setupPubDir "Docx2PDF_Install.exe") -Destination $rootInstaller -Force

# 6. Publish binary asset to GitHub Releases FIRST (before flipping version.json)
Write-Host "`n[6/7] Creating/Updating GitHub Release $tag..." -ForegroundColor Yellow
$existingRelease = $null
try {
    $existingRelease = Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/releases/tags/$tag"
} catch { }

$releaseBody = @"
# Docx2PDF $tag

$Notes

---

### Download (Single All-in-One Installer)

👉 **[Download Docx2PDF_Install.exe]($downloadUrl)**
"@

if ($null -ne $existingRelease) {
    $releaseId = $existingRelease.id
    Write-Host "Found existing release $tag (ID: $releaseId). Cleaning up old assets..."
    foreach ($asset in $existingRelease.assets) {
        Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/releases/assets/$($asset.id)" -Method Delete | Out-Null
    }
    Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/releases/$releaseId" -Method Patch -Body @{
        name = "Docx2PDF $tag"
        body = $releaseBody
    } | Out-Null
} else {
    Write-Host "Creating new GitHub release for $tag..."
    $newRel = Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/releases" -Method Post -Body @{
        tag_name   = $tag
        name       = "Docx2PDF $tag"
        body       = $releaseBody
        draft      = $false
        prerelease = $false
    }
    $releaseId = $newRel.id
}

Write-Host "Uploading $rootInstaller to release..."
$uploadUrl = "https://uploads.github.com/repos/$owner/$repo/releases/$releaseId/assets?name=Docx2PDF_Install.exe"
$uploadJson = & curl.exe -f -X POST -H "Authorization: Bearer $Token" -H "Content-Type: application/octet-stream" --data-binary "@$rootInstaller" "$uploadUrl"
if ($LASTEXITCODE -ne 0) { throw "Uploading Docx2PDF_Install.exe failed with exit code $LASTEXITCODE." }

# 7. Commit and push updated source code and version.json to GitHub main branch
Write-Host "`n[7/7] Pushing updated source code and version.json to GitHub main..." -ForegroundColor Yellow

$ref = Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/git/refs/heads/main"
$parentCommitSha = $ref.object.sha
$parentCommit = Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/git/commits/$parentCommitSha"
$remoteTree = Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/git/trees/$($parentCommit.tree.sha)?recursive=1"

$remoteShas = @{}
foreach ($entry in $remoteTree.tree) {
    if ($entry.type -eq "blob") {
        $remoteShas[$entry.path] = $entry.sha
    }
}

$excludedDirs = @('.git', '.vs', '.vscode', '.nuget', '.wix', '.dev', 'scratch', 'bin', 'obj', 'publish', 'setup_publish', 'TestResults')
$excludedExts = @('.exe', '.msi', '.wixpdb', '.pdb', '.tmp', '.log', '.token', '.user', '.suo', '.pfx', '.snk', '.key')
$excludedNames = @('.token', '.env')

$localFiles = Get-ChildItem -Path $rootDir -Recurse -File -Force | Where-Object {
    $rel = $_.FullName.Substring($rootDir.Length + 1) -replace '\\', '/'
    $segments = $rel.Split('/')
    $inExcludedDir = $false
    if ($segments.Length -gt 1) {
        foreach ($seg in $segments[0..($segments.Length - 2)]) {
            if ($seg -in $excludedDirs) { $inExcludedDir = $true; break }
        }
    }
    (-not $inExcludedDir) -and ($_.Extension -notin $excludedExts) -and ($_.Name -notin $excludedNames)
} | Sort-Object FullName

$treeItems = @()
$uploadedCount = 0

foreach ($file in $localFiles) {
    $relPath = $file.FullName.Substring($rootDir.Length + 1) -replace '\\', '/'
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $localSha = Get-GitBlobSha -Bytes $bytes

    if ($remoteShas.ContainsKey($relPath) -and $remoteShas[$relPath] -eq $localSha) {
        $blobSha = $localSha
    } else {
        Write-Host "Uploading changed/new file: $relPath..."
        $blobResp = Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/git/blobs" -Method Post -Body @{
            content  = [Convert]::ToBase64String($bytes)
            encoding = "base64"
        }
        $blobSha = $blobResp.sha
        $uploadedCount++
    }

    $treeItems += @{
        path = $relPath
        mode = "100644"
        type = "blob"
        sha  = $blobSha
    }
}

Write-Host "Creating Git tree ($($treeItems.Count) total files, $uploadedCount updated)..."
$treeResp = Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/git/trees" -Method Post -Body @{
    tree = $treeItems
}

$commitResp = Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/git/commits" -Method Post -Body @{
    message = "Release Docx2PDF $tag`n`n$Notes"
    tree    = $treeResp.sha
    parents = @($parentCommitSha)
}

$updateRefResp = Invoke-GitHubJson -Uri "https://api.github.com/repos/$owner/$repo/git/refs/heads/main" -Method Patch -Body @{
    sha   = $commitResp.sha
    force = $false
}

Write-Host "`n=======================================================" -ForegroundColor Green
Write-Host " SUCCESS! Docx2PDF $tag is live!" -ForegroundColor Green
Write-Host " Main branch commit: $($updateRefResp.object.sha)" -ForegroundColor Green
Write-Host " All running/installed apps will now auto-update." -ForegroundColor Green
Write-Host "=======================================================" -ForegroundColor Green
