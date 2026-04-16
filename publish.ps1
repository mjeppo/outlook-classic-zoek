param(
	[switch]$MakeInstaller
)

$ErrorActionPreference = "Stop"
$ISCC = "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
$ScriptDir = $PSScriptRoot

function Publish-Target {
	$outDir = "bin/Release/net8.0-windows/win-x64/publish"
	Write-Host "Publishing win-x64 to $outDir ..."
	dotnet publish -c Release -r win-x64 --self-contained false -p:PublishDir="$outDir/"
	if ($LASTEXITCODE -ne 0) { throw "dotnet publish mislukt (exit $LASTEXITCODE)" }
}

function Build-Installer {
	if (-not (Test-Path $ISCC)) {
		Write-Warning "ISCC.exe niet gevonden op: $ISCC -- installer overgeslagen."
		return
	}
	$fullIss = Join-Path $ScriptDir "installer-x64.iss"
	Write-Host "Installer bouwen: installer-x64.iss ..."
	& $ISCC $fullIss
	if ($LASTEXITCODE -ne 0) { throw "ISCC mislukt (exit $LASTEXITCODE)" }
}

Write-Host "Restoring packages..."
dotnet restore

Write-Host "Building release..."
dotnet build -c Release

Publish-Target
if ($MakeInstaller) { Build-Installer }

Write-Host "Klaar."
if ($MakeInstaller) {
	Write-Host "Installer staat in: $ScriptDir\installer-output\"
} else {
	Write-Host "Voeg -MakeInstaller toe om ook de Setup.exe te bouwen."
}
