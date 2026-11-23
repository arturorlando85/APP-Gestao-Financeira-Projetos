<#
  Equivalente ao build.py para máquinas sem Python.
    src/template.html + src/core.js -> dist/index.html (JSZip via CDN)
                                    -> dist/SafeHarbor_Financeiro.html (JSZip embutido)
                                    -> pacote-nova-pod/Financeiro/Gestao_Financeira_Projetos.html (nome neutro, para novas PODs)
  O JSZip vem de node_modules (npm install) ou, se não houver, do dist/SafeHarbor_Financeiro.html atual.
  Uso: powershell -ExecutionPolicy Bypass -File tools\build.ps1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$enc = New-Object Text.UTF8Encoding($false)
$read = { param($p) [IO.File]::ReadAllText((Join-Path $root $p), $enc) }
$write = { param($p, $t) [IO.File]::WriteAllText((Join-Path $root $p), $t, $enc); '{0}: {1:N0} KB' -f $p, ((Get-Item (Join-Path $root $p)).Length / 1KB) }

$CDN = '<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>'
$TITLE = '<title>Safe Harbor Financeiro</title>'
$template = & $read 'src\template.html'; $core = & $read 'src\core.js'
if (-not $template.Contains('/*__CORE__*/')) { throw 'marcador /*__CORE__*/ não encontrado no template' }
if (-not $template.Contains($CDN)) { throw 'tag do JSZip (CDN) não encontrada no template' }
$page = $template.Replace('/*__CORE__*/', $core)
& $write 'dist\index.html' $page

$nm = Join-Path $root 'node_modules\jszip\dist\jszip.min.js'
if (Test-Path $nm) { $jszip = '<script>' + "`n" + ([IO.File]::ReadAllText($nm, $enc).Replace('</script', '<\/script')) + "`n" + '</script>' }
else {
  $old = & $read 'dist\SafeHarbor_Financeiro.html'
  $m = [regex]::Match($old, '<script>\r?\n/\*!\r?\n\r?\nJSZip[\s\S]*?\r?\n</script>')
  if (-not $m.Success) { throw 'JSZip não encontrado: rode npm install ou mantenha o dist/SafeHarbor_Financeiro.html anterior' }
  $jszip = $m.Value
}
$local = $page.Replace($CDN, $jszip).Replace($TITLE, '<title>Safe Harbor Financeiro (local)</title>')
& $write 'dist\SafeHarbor_Financeiro.html' $local

# pacote para novas PODs: só os textos visíveis mudam (ids internos de trava/config continuam iguais)
$pod = $local.Replace('<title>Safe Harbor Financeiro (local)</title>', '<title>Gestão Financeira de Projetos</title>').
  Replace('<b>Safe Harbor</b>Controle financeiro · GPOD', '<b>Gestão Financeira</b>Controle financeiro de projetos').
  Replace('<h1>Safe Harbor · Controle financeiro de projetos</h1>', '<h1>Gestão Financeira · Controle financeiro de projetos</h1>')
if (Test-Path (Join-Path $root 'pacote-nova-pod\Financeiro')) { & $write 'pacote-nova-pod\Financeiro\Gestao_Financeira_Projetos.html' $pod }
