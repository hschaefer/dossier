# Convert .docx files to PDF with the installed desktop Word, so a Word export
# can be compared with Dossier's own PDF.
#
#   powershell -ExecutionPolicy Bypass -File .\docx-to-pdf.ps1 -In .\docx -Out .\word
#
# Get the folders from Dossier's sample matrix: sign in, then download
# /api/dev/word-samples and extract it. It holds docx/, chromium/ (the app's
# PDF for the same theme), this script and manifest.json.
#
# Every step is logged with a timestamp. If it stops at "open", Word is
# showing a dialog (recovery, activation, protected view): run it once with
# -Visible to see it, close Word and try again.
#
# Word's automation interface is only available in an interactive desktop
# session: run this from a normal PowerShell window, not as a service or a
# scheduled task, and close any open Word first. Run it from a local disk:
# Word's PDF export can hang when the documents or the output folder are on
# an SMB/network share.

[CmdletBinding()]
param(
  # A folder of .docx files, or one .docx file.
  [Parameter(Mandatory)] [string] $In,
  # Where the PDFs go; defaults to the input folder (the file's folder for one file).
  [string] $Out,
  [switch] $Force,
  # Show Word while it works, to reveal a dialog that blocks automation.
  [switch] $Visible
)

$ErrorActionPreference = "Stop"

function Write-Log([string] $Message) {
  Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message)
}

# Word is a single-threaded COM server. PowerShell 7 starts in an MTA, where
# the automation calls can hang; Windows PowerShell starts in an STA. Restart
# there instead of failing, so `pwsh -File` works too.
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne "STA") {
  $windowsPowerShell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
  if (-not (Test-Path -LiteralPath $windowsPowerShell)) {
    throw "Word automation needs a single-threaded apartment. Run this with Windows PowerShell: powershell.exe -STA -File .\docx-to-pdf.ps1 ..."
  }
  Write-Log "PowerShell is not in an STA; restarting in Windows PowerShell..."
  $arguments = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", ('"{0}"' -f $PSCommandPath), "-In", ('"{0}"' -f $In))
  if ($Out) { $arguments += @("-Out", ('"{0}"' -f $Out)) }
  if ($Force) { $arguments += "-Force" }
  if ($Visible) { $arguments += "-Visible" }
  $child = Start-Process -FilePath $windowsPowerShell -ArgumentList $arguments -NoNewWindow -Wait -PassThru
  exit $child.ExitCode
}

$source = Resolve-Path -LiteralPath $In
if (Test-Path -LiteralPath $source.Path -PathType Leaf) {
  $files = @(Get-Item -LiteralPath $source.Path)
  if (-not $Out) { $Out = $source.Path | Split-Path -Parent }
} else {
  $files = @(Get-ChildItem -LiteralPath $source.Path -Filter *.docx -File)
  if (-not $Out) { $Out = $source.Path }
}
$files = @($files | Where-Object { $_.Extension -ieq ".docx" })

if (-not (Test-Path -LiteralPath $Out)) {
  New-Item -ItemType Directory -Path $Out | Out-Null
}
$Out = (Resolve-Path -LiteralPath $Out).Path

if ($files.Count -eq 0) {
  Write-Log "No .docx files in $In."
  exit 1
}
Write-Log ("{0} document(s) -> {1}" -f $files.Count, $Out)

$running = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
if ($running.Count -gt 0) {
  Write-Warning "Word is already running. Close it, and dismiss any recovery pane, if this stops at 'open'."
}

Write-Log "Starting Word..."
$word = New-Object -ComObject Word.Application
$word.Visible = [bool]$Visible
$word.DisplayAlerts = 0          # wdAlertsNone
$word.AutomationSecurity = 3     # msoAutomationSecurityForceDisable: no macros
$word.Options.ConfirmConversions = $false
$word.Options.CheckSpellingAsYouType = $false
$word.Options.CheckGrammarAsYouType = $false
Write-Log ("Word {0} ready (visible: {1})" -f $word.Version, $word.Visible)

$done = 0
$failed = 0
try {
  foreach ($file in $files) {
    # Word's late-bound COM calls are safer with plain strings, so cast paths
    # explicitly before handing them over.
    $pdf = [string](Join-Path -Path $Out -ChildPath ($file.BaseName + ".pdf"))
    if ((Test-Path -LiteralPath $pdf) -and -not $Force) {
      Write-Log "skip  $($file.Name)"
      continue
    }

    Write-Log "open  $($file.Name)"
    $document = $null
    try {
      # A file downloaded from the internet opens in Protected View, where
      # automation cannot reach it.
      Unblock-File -LiteralPath $file.FullName -ErrorAction SilentlyContinue
      # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
      $document = $word.Documents.Open([string]$file.FullName, $false, $true, $false)
      Write-Log "opened $($file.Name)"
      Write-Log "save  $($file.Name)"
      try {
        $document.ExportAsFixedFormat([string]$pdf, 17)  # wdExportFormatPDF
      } catch {
        Write-Log "ExportAsFixedFormat failed ($($_.Exception.Message)); trying SaveAs2"
        $document.SaveAs2([string]$pdf, 17)              # wdFormatPDF
      }
      Write-Log "saved $($file.Name)"
      $document.Close($false)
      $document = $null
      $done++
      Write-Log "ok    $($file.Name) -> $pdf"
    } catch {
      $failed++
      if ($document) {
        try { $document.Close($false) } catch { }
      }
      Write-Warning "$($file.Name): $($_.Exception.Message)"
    }
  }
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

$missing = @($files | Where-Object { -not (Test-Path -LiteralPath (Join-Path $Out ($_.BaseName + ".pdf"))) })
Write-Host ""
Write-Log ("{0} converted, {1} failed, {2} missing. PDFs are in {3}" -f $done, $failed, $missing.Count, $Out)
if ($missing.Count -gt 0) {
  $missing | ForEach-Object { Write-Host "  missing: $($_.Name)" }
  exit 2
}

# Compare the two folders by eye or with any PDF tool:
#   chromium/<theme>.pdf  Dossier's Chromium output
#   word/<theme>.pdf      the same document through Microsoft Word
