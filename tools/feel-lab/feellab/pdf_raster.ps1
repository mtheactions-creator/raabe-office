# Renders the pages of a PDF to PNG files with Windows' own PDF renderer (Windows.Data.Pdf), so the look
# corpus (D-080 D) needs no extra software. Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File pdf_raster.ps1 <pdf> <outdir> <dpi>
# Writes <outdir>\page-1.png, page-2.png, ... and prints the page count.
param([string]$Pdf, [string]$OutDir, [int]$Dpi = 96)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq "IAsyncOperation``1" })[0]
$asAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq "AsTask" -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq "IAsyncAction" })[0]
function Await($op, [Type]$type) { $t = $asTask.MakeGenericMethod($type).Invoke($null, @($op)); $t.Wait(); $t.Result }
function AwaitAction($op) { $t = $asAction.Invoke($null, @($op)); $t.Wait() }
New-Item -ItemType Directory -Force $OutDir | Out-Null
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync((Resolve-Path $Pdf).Path)) ([Windows.Storage.StorageFile])
$doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
for ($i = 0; $i -lt $doc.PageCount; $i++) {
  $page = $doc.GetPage($i)
  $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
  $opts.DestinationWidth = [uint32]([math]::Round($page.Size.Width * $Dpi / 96))
  $opts.DestinationHeight = [uint32]([math]::Round($page.Size.Height * $Dpi / 96))
  $opts.BitmapEncoderId = [Windows.Graphics.Imaging.BitmapEncoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]::PngEncoderId
  $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
  AwaitAction ($page.RenderToStreamAsync($stream, $opts))
  $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream.GetInputStreamAt(0))
  $out = [System.IO.File]::Create((Join-Path $OutDir ("page-" + ($i + 1) + ".png")))
  $net.CopyTo($out); $out.Close(); $net.Close(); $page.Dispose()
}
Write-Output $doc.PageCount
