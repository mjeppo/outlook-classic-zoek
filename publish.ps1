param(
	[ValidateSet("x86", "x64", "both")]
	[string]$Architecture = "x64",
	[switch]$MakeInstaller
)

$ErrorActionPreference = "Stop"
$ISCC = "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
$ScriptDir = $PSScriptRoot

function Publish-Target([string]$Rid) {
	$outDir = "bin/Release/net8.0-windows/$Rid/publish"
	Write-Host "Publishing $Rid to $outDir ..."
	dotnet publish -c Release -r $Rid --self-contained false -p:PublishDir="$outDir/"
	if ($LASTEXITCODE -ne 0) { throw "dotnet publish mislukt voor $Rid (exit $LASTEXITCODE)" }
}

function Build-Installer([string]$IssFile) {
	if (-not (Test-Path $ISCC)) {
		Write-Warning "ISCC.exe niet gevonden op: $ISCC -- installer overgeslagen."
		return
	}
	$fullIss = Join-Path $ScriptDir $IssFile
	Write-Host "Installer bouwen: $IssFile ..."
	& $ISCC $fullIss
	if ($LASTEXITCODE -ne 0) { throw "ISCC mislukt voor $IssFile (exit $LASTEXITCODE)" }
}

Write-Host "Restoring packages..."
dotnet restore

Write-Host "Building release..."
dotnet build -c Release

switch ($Architecture) {
	"x86" {
		Publish-Target "win-x86"
		if ($MakeInstaller) { Build-Installer "installer-x86.iss" }
	}
	"x64" {
		Publish-Target "win-x64"
		if ($MakeInstaller) { Build-Installer "installer-x64.iss" }
	}
	"both" {
		Publish-Target "win-x86"
		Publish-Target "win-x64"
		if ($MakeInstaller) {
			Build-Installer "installer-x86.iss"
			Build-Installer "installer-x64.iss"
		}
	}
}

Write-Host "Klaar."
if ($MakeInstaller) {
	Write-Host "Installers staan in: $ScriptDir\installer-output\"
} else {
	Write-Host "Gebruik win-x86 als Outlook 32-bit is; win-x64 als Outlook 64-bit is."
	Write-Host "Voeg -MakeInstaller toe om ook de Setup.exe te bouwen."
}
