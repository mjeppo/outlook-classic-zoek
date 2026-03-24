param(
	[ValidateSet("x86", "x64", "both")]
	[string]$Architecture = "both"
)

$ErrorActionPreference = "Stop"

function Publish-Target([string]$Rid) {
	$outDir = "bin/Release/net8.0-windows/$Rid/publish"
	Write-Host "Publishing $Rid to $outDir ..."
	dotnet publish -c Release -r $Rid --self-contained true /p:PublishSingleFile=false -p:PublishDir="$outDir/"
}

Write-Host "Restoring packages..."
dotnet restore

Write-Host "Building release..."
dotnet build -c Release

switch ($Architecture) {
	"x86" { Publish-Target "win-x86" }
	"x64" { Publish-Target "win-x64" }
	"both" {
		Publish-Target "win-x86"
		Publish-Target "win-x64"
	}
}

Write-Host "Done. Use win-x86 if Outlook is 32-bit; use win-x64 if Outlook is 64-bit."
