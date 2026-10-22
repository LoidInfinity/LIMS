param(
    [Parameter(Mandatory=$true)][string]$Presentation,
    [Parameter(Mandatory=$true)][string]$Document,
    [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$presentationPath = (Resolve-Path -LiteralPath $Presentation).Path
$documentPath = (Resolve-Path -LiteralPath $Document).Path
$renderDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $renderDirectory) {
    throw "Output directory already exists: $renderDirectory"
}
New-Item -ItemType Directory -Path $renderDirectory | Out-Null
$powerPointApp = $null
$openedPresentation = $null
$wordApp = $null
$openedDocument = $null
try {
    $powerPointApp = New-Object -ComObject PowerPoint.Application
    # Open this generated fixture read-only, without a presentation window.
    $openedPresentation = $powerPointApp.Presentations.Open($presentationPath, -1, 0, 0)
    for ($slideIndex = 1; $slideIndex -le $openedPresentation.Slides.Count; $slideIndex++) {
        $pngPath = Join-Path $renderDirectory "slide-$slideIndex.png"
        $openedPresentation.Slides.Item($slideIndex).Export($pngPath, 'PNG', 1600, 900)
    }
    $openedPresentation.Close()
    $openedPresentation = $null
    $wordApp = New-Object -ComObject Word.Application
    $openedDocument = $wordApp.Documents.Open($documentPath, $false, $true, $false)
    $pdfPath = Join-Path $renderDirectory 'manuscript.pdf'
    $openedDocument.ExportAsFixedFormat($pdfPath, 17)
    $openedDocument.Close(0)
    $openedDocument = $null
    Get-ChildItem -LiteralPath $renderDirectory | Select-Object Name,Length
} finally {
    if ($null -ne $openedPresentation) { $openedPresentation.Close() }
    if ($null -ne $openedDocument) { $openedDocument.Close(0) }
    # Never close another document/presentation that belongs to the user.
    if ($null -ne $powerPointApp -and $powerPointApp.Presentations.Count -eq 0) { $powerPointApp.Quit() }
    if ($null -ne $wordApp -and $wordApp.Documents.Count -eq 0) { $wordApp.Quit() }
    foreach ($officeObject in @($powerPointApp, $wordApp)) {
        if ($null -ne $officeObject) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($officeObject) }
    }
}
