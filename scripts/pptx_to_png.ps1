# 将 pptx 每页导出为 PNG（PowerPoint COM；仅本机验证用，不属于服务管线）
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$OutDir = "."
)
$ErrorActionPreference = "Stop"
$Path = (Resolve-Path $Path).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$base = [IO.Path]::GetFileNameWithoutExtension($Path)
$pp = New-Object -ComObject PowerPoint.Application
try {
    $pres = $pp.Presentations.Open($Path, $true, $false, $false)
    $i = 0
    foreach ($slide in $pres.Slides) {
        $i++
        $out = Join-Path $OutDir ("{0}_p{1}.png" -f $base, $i)
        $slide.Export($out, "PNG", 1280, 720) | Out-Null
        Write-Output "exported: $out"
    }
    $pres.Close()
}
finally { $pp.Quit() }
