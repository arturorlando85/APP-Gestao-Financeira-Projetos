<#
  Gera Database.xlsx e Config.xlsx ZERADOS para uma nova POD, na estrutura que o app lê
  (src/core.js: load / loadConfig). Sem Excel: monta o pacote OOXML direto.

  Uso:  powershell -ExecutionPolicy Bypass -File tools\gerar_planilhas_modelo.ps1 -Saida "pacote-nova-pod\Financeiro"
#>
param([string]$Saida = "pacote-nova-pod\Financeiro", [int]$AnoDiasUteis = 2027)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8 = New-Object System.Text.UTF8Encoding($false)

# ---------------- utilidades XML ----------------
function Col([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [math]::Floor(($n - 1) / 26) }; $s }
function Esc([string]$t) { $t.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;') }
function Cell([string]$ref, $v, [int]$s = 0) {
  $sa = if ($s) { " s=`"$s`"" } else { '' }
  if ($null -eq $v -or "$v" -eq '') { return "<c r=`"$ref`"$sa/>" }
  if ($v -is [double] -or $v -is [int] -or $v -is [decimal]) { return "<c r=`"$ref`"$sa><v>$([string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0}', $v))</v></c>" }
  return "<c r=`"$ref`"$sa t=`"inlineStr`"><is><t xml:space=`"preserve`">$(Esc $v)</t></is></c>"
}
# linhas: hashtable  rowNumber -> array de @(colNumber, valor, estilo)
function SheetXml($rows, [string]$dim, [string]$cols = '', [string]$pane = '', [bool]$table = $false) {
  $sb = New-Object System.Text.StringBuilder
  foreach ($r in ($rows.Keys | Sort-Object { [int]$_ })) {
    [void]$sb.Append("<row r=`"$r`">")
    $cells = @($rows[$r] | Sort-Object { $_[0] })   # @() evita que uma linha com uma única célula seja "desembrulhada"
    foreach ($c in $cells) { [void]$sb.Append((Cell "$(Col $c[0])$r" $c[1] $c[2])) }
    [void]$sb.Append('</row>')
  }
  $sv = if ($pane) { "<sheetView workbookViewId=`"0`">$pane</sheetView>" } else { '<sheetView workbookViewId="0"/>' }
  $tp = if ($table) { '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' } else { '' }
  "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><dimension ref=`"$dim`"/><sheetViews>$sv</sheetViews><sheetFormatPr defaultRowHeight=`"15`"/>$(if ($cols) { "<cols>$cols</cols>" })<sheetData>$($sb.ToString())</sheetData>$tp</worksheet>"
}
function ColsXml([int[]]$widths) { $i = 0; ($widths | ForEach-Object { $i++; "<col min=`"$i`" max=`"$i`" width=`"$_`" customWidth=`"1`"/>" }) -join '' }
function TableXml([int]$id, [string]$name, [string]$ref, [string[]]$headers) {
  $tc = ''; $i = 0; foreach ($h in $headers) { $i++; $tc += "<tableColumn id=`"$i`" name=`"$(Esc $h)`"/>" }
  "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<table xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" id=`"$id`" name=`"$name`" displayName=`"$name`" ref=`"$ref`" totalsRowShown=`"0`"><autoFilter ref=`"$ref`"/><tableColumns count=`"$($headers.Count)`">$tc</tableColumns><tableStyleInfo name=`"TableStyleLight1`" showFirstColumn=`"0`" showLastColumn=`"0`" showRowStripes=`"1`" showColumnStripes=`"0`"/></table>"
}

# estilos: monocromático AI/R (cabeçalho carbono #222222 com texto branco)
$STYLES = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="3"><numFmt numFmtId="164" formatCode="&quot;R$&quot; #,##0.00"/><numFmt numFmtId="165" formatCode="mmm/yyyy"/><numFmt numFmtId="166" formatCode="dd/mm/yyyy"/></numFmts>
<fonts count="4"><font><sz val="10"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font><font><sz val="16"/><name val="Arial"/><family val="2"/></font><font><i/><sz val="10"/><color rgb="FF595959"/><name val="Arial"/><family val="2"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF222222"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE9E9E9"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="10">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
'@
$S_HEAD = 1; $S_TITLE = 2; $S_BRL = 3; $S_MES = 4; $S_DATA = 5; $S_PCT = 6; $S_NUM = 7; $S_NOTE = 8; $S_BOX = 9

# escreve o .xlsx: $sheets = lista ordenada de @{ name; xml; table = @{ xml } (opcional) }
function Write-Xlsx([string]$path, $sheets) {
  if (Test-Path $path) { Remove-Item $path -Force }
  $fs = [IO.File]::Open($path, 'Create'); $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
  $put = { param($name, $text) $e = $zip.CreateEntry($name, [IO.Compression.CompressionLevel]::Optimal); $w = New-Object IO.StreamWriter($e.Open(), $utf8); $w.Write($text); $w.Close() }
  $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
  $wbSheets = ''; $wbRels = ''; $i = 0; $t = 0
  foreach ($s in $sheets) {
    $i++
    & $put "xl/worksheets/sheet$i.xml" $s.xml
    $ct += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
    $wbSheets += "<sheet name=`"$(Esc $s.name)`" sheetId=`"$i`" r:id=`"rId$i`"/>"
    $wbRels += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
    if ($s.table) {
      $t++
      & $put "xl/tables/table$t.xml" $s.table
      $ct += "<Override PartName=`"/xl/tables/table$t.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml`"/>"
      & $put "xl/worksheets/_rels/sheet$i.xml.rels" ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + "<Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table`" Target=`"../tables/table$t.xml`"/></Relationships>")
    }
  }
  $n = $i + 1
  $wbRels += "<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>"
  & $put '[Content_Types].xml' ($ct + '</Types>')
  & $put '_rels/.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')
  & $put 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + "<workbook xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><bookViews><workbookView xWindow=`"0`" yWindow=`"0`" windowWidth=`"28800`" windowHeight=`"15000`"/></bookViews><sheets>$wbSheets</sheets><calcPr calcId=`"191029`" fullCalcOnLoad=`"1`"/></workbook>")
  & $put 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + "<Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$wbRels</Relationships>")
  & $put 'xl/styles.xml' $STYLES
  $now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
  & $put 'docProps/core.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + "<cp:coreProperties xmlns:cp=`"http://schemas.openxmlformats.org/package/2006/metadata/core-properties`" xmlns:dc=`"http://purl.org/dc/elements/1.1/`" xmlns:dcterms=`"http://purl.org/dc/terms/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`"><dc:title>Gestão Financeira de Projetos</dc:title><dcterms:created xsi:type=`"dcterms:W3CDTF`">$now</dcterms:created></cp:coreProperties>")
  & $put 'docProps/app.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>')
  $zip.Dispose(); $fs.Close()
}
function Add($rows, [int]$r, [int]$c, $v, [int]$s = 0) { if (-not $rows.ContainsKey($r)) { $rows[$r] = New-Object System.Collections.ArrayList }; [void]$rows[$r].Add(@($c, $v, $s)) }

# ---------------- dados de referência (genéricos, não específicos de uma POD) ----------------
$REVENUE = @('CONFIRMADO', 'FECHADO', 'CHANCE ALTA', 'CHANCE MÉDIA', 'CHANCE BAIXA', 'PIPELINE', 'CANCELADO')
$SENIOR = @('Trainee', 'Júnior', 'Pleno', 'Sênior', 'Especialista')
$TAX = [ordered]@{ PIS = 0.0065; COFINS = 0.03; INSS = 0.027; ISSQN = 0.02; OPEX = 0.0981 }
$MESES = @('janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro')

# dias úteis do ano (feriados nacionais, incluindo Carnaval, Sexta-feira Santa e Corpus Christi)
function Easter([int]$y) { $a = $y % 19; $b = [math]::Floor($y / 100); $c = $y % 100; $d = [math]::Floor($b / 4); $e = $b % 4; $f = [math]::Floor(($b + 8) / 25); $g = [math]::Floor(($b - $f + 1) / 3); $h = (19 * $a + $b - $d - $g + 15) % 30; $i = [math]::Floor($c / 4); $k = $c % 4; $l = (32 + 2 * $e + 2 * $i - $h - $k) % 7; $m = [math]::Floor(($a + 11 * $h + 22 * $l) / 451); $mo = [math]::Floor(($h + $l - 7 * $m + 114) / 31); $dy = (($h + $l - 7 * $m + 114) % 31) + 1; Get-Date -Year $y -Month $mo -Day $dy -Hour 0 -Minute 0 -Second 0 -Millisecond 0 }
$E = Easter $AnoDiasUteis
$FERIADOS = @("$AnoDiasUteis-01-01", "$AnoDiasUteis-04-21", "$AnoDiasUteis-05-01", "$AnoDiasUteis-09-07", "$AnoDiasUteis-10-12", "$AnoDiasUteis-11-02", "$AnoDiasUteis-11-15", "$AnoDiasUteis-11-20", "$AnoDiasUteis-12-25") +
  @($E.AddDays(-48), $E.AddDays(-47), $E.AddDays(-2), $E.AddDays(60) | ForEach-Object { $_.ToString('yyyy-MM-dd') })
$DIAS = @{}
for ($mo = 1; $mo -le 12; $mo++) {
  $n = 0; $d = Get-Date -Year $AnoDiasUteis -Month $mo -Day 1 -Hour 0 -Minute 0 -Second 0
  while ($d.Month -eq $mo) { if ($d.DayOfWeek -ne 'Saturday' -and $d.DayOfWeek -ne 'Sunday' -and $FERIADOS -notcontains $d.ToString('yyyy-MM-dd')) { $n++ }; $d = $d.AddDays(1) }
  $DIAS[$mo] = $n
}

New-Item -ItemType Directory -Force -Path $Saida | Out-Null
$Saida = (Resolve-Path $Saida).Path

# ======================= Database.xlsx =======================
# 1. Config (espelho técnico que as fórmulas das abas BL_ usam)
$r = @{}
Add $r 1 1 'Status da Proposta' $S_HEAD; Add $r 1 3 'Revenue State' $S_HEAD; Add $r 1 5 'iStudio' $S_HEAD; Add $r 1 7 'Perfil' $S_HEAD; Add $r 1 9 'Senioridade' $S_HEAD
Add $r 1 15 'Clientes' $S_HEAD; Add $r 1 16 'BP' $S_HEAD; Add $r 1 18 'PowerHouse' $S_HEAD
for ($i = 0; $i -lt $REVENUE.Count; $i++) { Add $r ($i + 2) 3 $REVENUE[$i] }
for ($i = 0; $i -lt $SENIOR.Count; $i++) { Add $r ($i + 2) 9 $SENIOR[$i] }
# impostos: cabeçalho na linha 2, alíquotas na linha 3 (U:Y) -> as fórmulas BL leem '1. Config'!U$3
$c = 21; foreach ($k in $TAX.Keys) { Add $r 2 $c $k $S_HEAD; Add $r 3 $c ([double]$TAX[$k]) $S_PCT; $c++ }
# dias úteis: T:V a partir da linha 6
Add $r 5 20 'Mês' $S_HEAD; Add $r 5 21 'Dias úteis' $S_HEAD; Add $r 5 22 'Horas úteis' $S_HEAD
for ($mo = 1; $mo -le 12; $mo++) { Add $r ($mo + 5) 20 $MESES[$mo - 1]; Add $r ($mo + 5) 21 ([int]$DIAS[$mo]); Add $r ($mo + 5) 22 ([int]$DIAS[$mo] * 8) }
# clientes x tabelas (AA:AE) e chaves de tributação (AF:AG)
Add $r 1 27 'Cliente' $S_HEAD; Add $r 1 28 'Aba' $S_HEAD; Add $r 1 29 'Tabela' $S_HEAD; Add $r 1 30 'BP' $S_HEAD; Add $r 1 31 'Status' $S_HEAD
Add $r 1 32 'Chave (Cliente|Projeto)' $S_HEAD; Add $r 1 33 'Tributação' $S_HEAD
$cfgSheet = SheetXml $r 'A1:AG17' (ColsXml @(18, 2, 16, 2, 16, 2, 18, 2, 14, 2, 2, 2, 2, 2, 28, 16, 2, 18, 2, 14, 11, 11, 11, 11, 11, 2, 28, 22, 22, 14, 10, 40, 12))

# 2. Propostas (contratos)
$PROP = @('PowerHouse', 'Proposta', 'Cliente', 'Projeto', 'Ano Fiscal', 'Chance de Conversão', 'iStudio', 'Pre Sales Lead', 'Start Date', 'End Date', 'COLT OPT', 'Valor Assinado', 'Valor Usado', 'Delta', 'Funding', 'Projeto - Config')
$PROPS = @{ 'Start Date' = $S_DATA; 'End Date' = $S_DATA; 'Valor Assinado' = $S_BRL; 'Valor Usado' = $S_BRL; 'Delta' = $S_BRL }
$r = @{}; for ($i = 0; $i -lt $PROP.Count; $i++) { $h = $PROP[$i]; Add $r 1 ($i + 1) $h $S_HEAD; Add $r 2 ($i + 1) $null ([int]$PROPS[$h]) }
$lastP = Col $PROP.Count
$propSheet = SheetXml $r "A1:${lastP}2" (ColsXml @(16, 30, 26, 30, 10, 18, 18, 20, 12, 12, 12, 16, 16, 16, 10, 40)) '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' $true
$propTable = TableXml 1 'tbl_Propostas' "A1:${lastP}2" $PROP

# BL_Template (modelo das abas de billing por cliente; o app cria BL_<Cliente> a partir dela)
$BL = @('Nome da Proposta', 'Cliente', 'Projeto', 'iStudio', 'Papel', 'Senioridade', 'Vaga Intranet', 'Profissional', 'Custo Profissional', 'R$/H', 'Mês de Execução', 'Horas Provisionadas', 'Horas Executadas', 'Valor Bruto', 'Valor Liquido', 'Resultado Bruto', 'MB Profissional', 'Status da Receita', 'Obs', 'PIS', 'COFINS', 'INSS', 'ISSQN', 'Custo Total Profissional')
$BLS = @{ 'Custo Profissional' = $S_BRL; 'R$/H' = $S_BRL; 'Mês de Execução' = $S_MES; 'Horas Provisionadas' = $S_NUM; 'Horas Executadas' = $S_NUM; 'Valor Bruto' = $S_BRL; 'Valor Liquido' = $S_BRL; 'Resultado Bruto' = $S_BRL; 'MB Profissional' = $S_PCT; 'PIS' = $S_BRL; 'COFINS' = $S_BRL; 'INSS' = $S_BRL; 'ISSQN' = $S_BRL; 'Custo Total Profissional' = $S_BRL }
$r = @{}; Add $r 1 1 'Billing — Template' $S_TITLE
for ($i = 0; $i -lt $BL.Count; $i++) { $h = $BL[$i]; Add $r 3 ($i + 1) $h $S_HEAD; Add $r 4 ($i + 1) $null ([int]$BLS[$h]) }
$lastB = Col $BL.Count
$blSheet = SheetXml $r "A1:${lastB}4" (ColsXml @(26, 22, 26, 16, 22, 13, 13, 26, 14, 12, 12, 12, 12, 14, 14, 14, 10, 16, 30, 12, 12, 12, 12, 16)) '<pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>' $true
$blTable = TableXml 2 'tbl_Template' "A3:${lastB}4" $BL

Write-Xlsx (Join-Path $Saida 'Database.xlsx') @(
  @{ name = '1. Config'; xml = $cfgSheet },
  @{ name = '2. Propostas'; xml = $propSheet; table = $propTable },
  @{ name = 'BL_Template'; xml = $blSheet; table = $blTable }
)

# ======================= Config.xlsx =======================
function ListSheet([string]$header, $values, [int]$width = 28) {
  $r = @{}; Add $r 1 1 $header $S_HEAD; $i = 2; foreach ($v in $values) { Add $r $i 1 $v; $i++ }
  SheetXml $r "A1:A$([math]::Max(1, $i - 1))" (ColsXml @($width)) '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
}
function GridSheet([string[]]$headers, $data, [int[]]$widths, $styles = @{}) {
  $r = @{}; for ($j = 0; $j -lt $headers.Count; $j++) { Add $r 1 ($j + 1) $headers[$j] $S_HEAD }
  $i = 2; foreach ($row in $data) { for ($j = 0; $j -lt $row.Count; $j++) { Add $r $i ($j + 1) $row[$j] ([int]$styles[$j]) }; $i++ }
  SheetXml $r "A1:$(Col $headers.Count)$([math]::Max(1, $i - 1))" (ColsXml $widths) '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
}

$LEIA = @(
  @('Aba', 'O que preencher', 'Quem mantém'),
  @('Impostos', 'Alíquotas de PIS, COFINS, INSS, ISSQN e o OPEX (overhead aplicado ao custo). Valores iniciais são os padrões do app: confirme com o Financeiro.', 'Operations / Business Manager'),
  @('Perfis', 'OPCIONAL: só é usada se não houver a planilha DRP na pasta Financeiro. Innovation Studios, perfis, custos e ratecards vêm do DRP (aba Custo Referência), fornecido pela gestão.', 'Operations / Business Manager'),
  @('Dias Uteis', "Dias e horas úteis de cada mês ($AnoDiasUteis, feriados nacionais). Ajuste feriados locais.", 'Operations Manager'),
  @('Revenue State', 'Status usados em "Chance de conversão" do contrato e em "Status da receita" da alocação.', 'Operations Manager'),
  @('Status da Proposta', 'Lista de status de proposta (opcional).', 'Operations Manager'),
  @('PowerHouse', 'Lista de PowerHouses que aparecem no cadastro de contratos.', 'Business Manager'),
  @('Projetos', 'NÃO EDITAR: o app mantém (uma linha por contrato, com tributação, tipo de projeto e Margem DRP informados ao criar ou editar o projeto).', 'App'),
  @('Clientes x Tabelas', 'Clientes disponíveis no app. Pode cadastrar clientes aqui (só a coluna Cliente); o app preenche Aba e Tabela quando cria a aba de billing.', 'App / Business Manager'),
  @('BP por Cliente', 'Business Partner responsável por cada cliente.', 'Business Manager')
)
$r = @{}; Add $r 1 1 'Config · Gestão Financeira de Projetos' $S_TITLE
Add $r 2 1 'Regras do app: impostos, dias úteis e listas. Innovation Studios, perfis, custos e ratecards vêm da planilha DRP (da gestão), que deve ficar na mesma pasta Financeiro.' $S_NOTE
$i = 4; foreach ($row in $LEIA) { $st = if ($i -eq 4) { $S_HEAD } else { $S_BOX }; for ($j = 0; $j -lt 3; $j++) { Add $r $i ($j + 1) $row[$j] $st }; $i++ }
$leiaSheet = SheetXml $r "A1:C$($i - 1)" (ColsXml @(22, 90, 30))

$taxData = @(, @($TAX.Values | ForEach-Object { [double]$_ }))
$diasData = for ($mo = 1; $mo -le 12; $mo++) { , @($MESES[$mo - 1], [int]$DIAS[$mo], [int]$DIAS[$mo] * 8) }

Write-Xlsx (Join-Path $Saida 'Config.xlsx') @(
  @{ name = 'Leia-me'; xml = $leiaSheet },
  @{ name = 'Impostos'; xml = (GridSheet @('PIS', 'COFINS', 'INSS', 'ISSQN', 'OPEX') $taxData @(12, 12, 12, 12, 12) @{ 0 = $S_PCT; 1 = $S_PCT; 2 = $S_PCT; 3 = $S_PCT; 4 = $S_PCT }) },
  @{ name = 'Perfis'; xml = (GridSheet @('iStudio', 'Perfil', 'Senioridade', 'Custo Total 2026', 'Ratecard 2026', 'Custo Total 2027 PC', 'Ratecard 2027') @() @(24, 30, 14, 18, 16, 20, 16)) },
  @{ name = 'Dias Uteis'; xml = (GridSheet @('Mês', 'Dias úteis', 'Horas úteis') $diasData @(14, 12, 12)) },
  @{ name = 'Revenue State'; xml = (ListSheet 'Revenue State' $REVENUE) },
  @{ name = 'Status da Proposta'; xml = (ListSheet 'Status da Proposta' @()) },
  @{ name = 'PowerHouse'; xml = (ListSheet 'PowerHouse' @()) },
  @{ name = 'Projetos'; xml = (GridSheet @('Status', 'Chave (Cliente|Projeto)', 'Tributação', 'Tipo de Projeto', 'Margem DRP') @() @(12, 50, 14, 20, 14)) },
  @{ name = 'Clientes x Tabelas'; xml = (GridSheet @('Cliente', 'Aba', 'Tabela') @() @(30, 30, 30)) },
  @{ name = 'BP por Cliente'; xml = (GridSheet @('Cliente', 'BP') @() @(30, 30)) }
)

"Dias úteis $AnoDiasUteis : " + (($DIAS.Keys | Sort-Object | ForEach-Object { "$($MESES[$_-1].Substring(0,3))=$($DIAS[$_])" }) -join ' ')
Get-ChildItem $Saida | Select-Object Name, Length | Format-Table -AutoSize | Out-String
