#!/usr/bin/env pwsh
# Render every slide of a PPTX to PNG through PowerPoint COM for visual QA.
param(
    [Parameter(Mandatory = $true)][string]$Presentation,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [int]$Width = 1920,
    [int]$Height = 1080
)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $Presentation).Path
$target = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $target | Out-Null
if (Get-ChildItem -LiteralPath $target -Filter '*.png' | Select-Object -First 1) {
    throw "Output directory already contains PNG files: $target"
}
$powerpoint = New-Object -ComObject PowerPoint.Application
try {
    $document = $powerpoint.Presentations.Open($source, $true, $false, $false)
    try {
        for ($i = 1; $i -le $document.Slides.Count; $i++) {
            $file = Join-Path $target ('{0:D2}.png' -f $i)
            $document.Slides.Item($i).Export($file, 'PNG', $Width, $Height)
        }
        Write-Output ("Rendered {0} slide(s) -> {1}" -f $document.Slides.Count, $target)
    }
    finally {
        $document.Close()
    }
}
finally {
    $powerpoint.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
}
