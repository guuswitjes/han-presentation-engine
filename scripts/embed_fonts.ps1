<#
.SYNOPSIS
    Sluit lettertypen in een .pptx in, zodat de deck er ook goed uitziet op een pc
    zonder Roboto / Roboto Condensed.

.WHY
    python-pptx kan geen fonts insluiten. De HAN-template verwijst alleen naar
    'Roboto Condensed SemiBold' en 'Roboto' bij naam. Staat dat font niet op de
    ontvangende pc, dan valt PowerPoint terug op een standaardfont en klopt de
    huisstijl niet meer. Dit script opent de deck via PowerPoint zelf (COM),
    zet 'lettertypen insluiten' aan, en slaat opnieuw op.

    Roboto is een vrij font (Apache License) en mag ingesloten worden.

.REQUIREMENTS
    - Microsoft PowerPoint geïnstalleerd (gevonden: Office16).
    - Roboto + Roboto Condensed geïnstalleerd op DEZE pc (alleen geïnstalleerde
      fonts kunnen worden ingesloten). Op Guus' pc is dat het geval.

.USAGE
    powershell -ExecutionPolicy Bypass -File embed_fonts.ps1 -Path "pad\deck.pptx"
    # optioneel: -OnlyUsed  -> alleen gebruikte tekens insluiten (kleiner bestand)

.NOTE
    Draai dit vanuit een eigen PowerShell-venster als Claude Code's sandbox de
    COM-call blokkeert. Het is een naverwerkstap NA het python-pptx bouwen.
#>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [switch]$OnlyUsed
)

$ErrorActionPreference = "Stop"

$full = (Resolve-Path -LiteralPath $Path).Path
if (-not (Test-Path -LiteralPath $full)) { throw "Bestand niet gevonden: $full" }

Write-Host "Fonts insluiten in: $full"

# PowerPoint-constanten
$ppSaveAsOpenXMLPresentation = 24   # .pptx
$msoTrue = -1
$msoFalse = 0

$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($full, $false, $false, $false)  # ReadOnly=false, Untitled=false, WithWindow=false

    # Subset-instelling vooraf zetten (alleen gebruikte tekens = kleiner bestand)
    try { $pres.SaveSubsetFonts = [bool]$OnlyUsed } catch { }

    # Insluiten gebeurt via de derde parameter van SaveAs (EmbedTrueTypeFonts).
    # De property .EmbedTrueTypeFonts bestaat niet in elke versie; SaveAs wel.
    $pres.SaveAs($full, $ppSaveAsOpenXMLPresentation, $msoTrue)

    Write-Host "Klaar. Fonts ingesloten (subset: $([bool]$OnlyUsed))."
}
finally {
    if ($pres) { $pres.Close() | Out-Null }
    if ($ppt)  { $ppt.Quit()  | Out-Null }
    if ($pres) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null }
    if ($ppt)  { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)  | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
