[CmdletBinding()]
param(
  [string]$WorkbookPath = 'C:\Users\user\Desktop\CONTROLE DE PRAZOS 2026.xlsx',
  [string]$OutputPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Normalize-PlainText {
  param(
    [AllowNull()]
    [string]$Value
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return ''
  }

  return (($Value -replace '\s+', ' ').Trim())
}

function Convert-ToLookupKey {
  param(
    [AllowNull()]
    [string]$Value
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return ''
  }

  $normalized = $Value.Normalize([System.Text.NormalizationForm]::FormD)
  $chars = New-Object System.Collections.Generic.List[char]

  foreach ($char in $normalized.ToCharArray()) {
    $category = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
    if ($category -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$chars.Add($char)
    }
  }

  return (-join $chars.ToArray()).ToUpperInvariant().Trim()
}

function Get-SheetConfig {
  param(
    [string]$SheetName
  )

  switch (Convert-ToLookupKey $SheetName) {
    'PRORROGACAO' {
      return [ordered]@{
        Category = 'PRORROGAÇÃO'
        StartRow = 3
        MaxColumns = 11
        Mode = 'Default'
        Mapping = [ordered]@{
          Modalidade = 3
          Objeto = 4
          Processo = 5
          Numero = 6
          Valor = 7
          Fornecedor = 8
          Inicio = 9
          Fim = 10
          Observacoes = 11
          Gestor = 0
        }
      }
    }
    'LOCACAO' {
      return [ordered]@{
        Category = 'LOCAÇÃO'
        StartRow = 3
        MaxColumns = 11
        Mode = 'Default'
        Mapping = [ordered]@{
          Modalidade = 1
          Objeto = 2
          Processo = 3
          Numero = 4
          Valor = 5
          Fornecedor = 6
          Inicio = 7
          Fim = 8
          Observacoes = 9
          Gestor = 11
        }
      }
    }
    'CONCORRENCIA ELETRONICA' {
      return [ordered]@{
        Category = 'CONCORRÊNCIA ELETRÔNICA'
        StartRow = 3
        MaxColumns = 11
        Mode = 'Default'
        Mapping = [ordered]@{
          Modalidade = 1
          Objeto = 2
          Processo = 3
          Numero = 4
          Valor = 5
          Fornecedor = 6
          Inicio = 7
          Fim = 8
          Observacoes = 9
          Gestor = 10
        }
      }
    }
    'CONCORRENCIA PRESENCIAL' {
      return [ordered]@{
        Category = 'CONCORRÊNCIA PRESENCIAL'
        StartRow = 3
        MaxColumns = 11
        Mode = 'Default'
        Mapping = [ordered]@{
          Modalidade = 1
          Objeto = 2
          Processo = 3
          Numero = 4
          Valor = 5
          Fornecedor = 6
          Inicio = 7
          Fim = 8
          Observacoes = 9
          Gestor = 10
        }
      }
    }
    'TOMADA DE PRECOS' {
      return [ordered]@{
        Category = 'TOMADA DE PREÇOS'
        StartRow = 3
        MaxColumns = 11
        Mode = 'Default'
        Mapping = [ordered]@{
          Modalidade = 1
          Objeto = 2
          Processo = 3
          Numero = 4
          Valor = 5
          Fornecedor = 6
          Inicio = 7
          Fim = 8
          Observacoes = 9
          Gestor = 10
        }
      }
    }
    'CHAMADA PUBLICA' {
      return [ordered]@{
        Category = 'CHAMADA PÚBLICA'
        StartRow = 3
        MaxColumns = 11
        Mode = 'Default'
        Mapping = [ordered]@{
          Modalidade = 1
          Objeto = 2
          Processo = 3
          Numero = 4
          Fornecedor = 5
          Valor = 6
          Inicio = 7
          Fim = 8
          Observacoes = 9
          Gestor = 10
        }
      }
    }
    'PREGAO ELETRONICO' {
      return [ordered]@{
        Category = 'PREGÃO ELETRÔNICO'
        StartRow = 3
        MaxColumns = 11
        Mode = 'Pregao'
        Mapping = [ordered]@{
          Modalidade = 1
          Objeto = 2
          Processo = 3
          Numero = 4
          Inicio = 7
          Fim = 8
          Observacoes = 9
          Gestor = 10
        }
      }
    }
    'DISPENSA' {
      return [ordered]@{
        Category = 'DISPENSA'
        StartRow = 2
        MaxColumns = 8
        Mode = 'Default'
        Mapping = [ordered]@{
          Modalidade = 1
          Objeto = 2
          Processo = 3
          Numero = 4
          Valor = 5
          Fornecedor = 6
          Inicio = 7
          Fim = 8
          Observacoes = 0
          Gestor = 0
        }
      }
    }
    default {
      return $null
    }
  }
}

function Get-CellText {
  param(
    $Worksheet,
    [int]$Row,
    [int]$Column
  )

  if ($Column -le 0) {
    return ''
  }

  return Normalize-PlainText ([string]$Worksheet.Cells.Item($Row, $Column).Text)
}

function Get-RowCells {
  param(
    $Worksheet,
    [int]$Row,
    [int]$MaxColumns
  )

  $cells = @()
  for ($column = 1; $column -le $MaxColumns; $column++) {
    $cells += Get-CellText -Worksheet $Worksheet -Row $Row -Column $column
  }

  return $cells
}

function Test-RowHasData {
  param(
    [string[]]$Cells
  )

  foreach ($cell in $Cells) {
    if (-not [string]::IsNullOrWhiteSpace($cell)) {
      return $true
    }
  }

  return $false
}

function Get-TextByIndex {
  param(
    [string[]]$Cells,
    [int]$Index
  )

  if ($Index -le 0) {
    return ''
  }

  $zeroBased = $Index - 1
  if ($zeroBased -ge $Cells.Count) {
    return ''
  }

  return Normalize-PlainText $Cells[$zeroBased]
}

function Convert-DateToIso {
  param(
    [AllowNull()]
    [string]$Text
  )

  $value = Normalize-PlainText $Text
  if (-not $value) {
    return ''
  }

  $match = [regex]::Match($value, '(\d{2}/\d{2}/\d{4})')
  if (-not $match.Success) {
    return ''
  }

  $parsed = [datetime]::MinValue
  $ok = [datetime]::TryParseExact(
    $match.Groups[1].Value,
    'dd/MM/yyyy',
    [System.Globalization.CultureInfo]::InvariantCulture,
    [System.Globalization.DateTimeStyles]::None,
    [ref]$parsed
  )

  if (-not $ok) {
    return ''
  }

  return $parsed.ToString('yyyy-MM-dd')
}

function Get-AnoFromNumero {
  param(
    [AllowNull()]
    [string]$Numero
  )

  $match = [regex]::Match((Normalize-PlainText $Numero), '(19|20)\d{2}')
  if ($match.Success) {
    return [int]$match.Value
  }

  return ''
}

function Get-TipoFromNumero {
  param(
    [AllowNull()]
    [string]$Numero
  )

  if ((Normalize-PlainText $Numero) -match '^ATA\b') {
    return 'Ata'
  }

  return 'Contrato'
}

function Test-LooksLikeStatus {
  param(
    [AllowNull()]
    [string]$Text
  )

  $value = Convert-ToLookupKey (Normalize-PlainText $Text)
  if (-not $value) {
    return $false
  }

  return ($value -match 'SUSPENSO|FRACASSADO|EM ANDAMENTO|ANDAMENTO|LICITANDO|ENCERRADO|FINALIZADO')
}

function Get-StatusFromCells {
  param(
    [string[]]$Cells
  )

  $joined = Convert-ToLookupKey (($Cells | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' | ')

  if ($joined -match 'SUSPENSO') {
    return 'SUSPENSO'
  }

  if ($joined -match 'FRACASSADO') {
    return 'FRACASSADO'
  }

  if ($joined -match 'EM ANDAMENTO|ANDAMENTO|LICITANDO') {
    return 'EM ANDAMENTO'
  }

  if ($joined -match 'ENCERRADO|FINALIZADO') {
    return 'ENCERRADO'
  }

  return 'VIGENTE'
}

function Convert-ContractValue {
  param(
    [AllowNull()]
    [string]$RawText
  )

  $value = Normalize-PlainText $RawText
  if (-not $value) {
    return [ordered]@{
      Value = 0.0
      ValueText = ''
      IsNumeric = $false
    }
  }

  if (Test-LooksLikeStatus $value) {
    return [ordered]@{
      Value = 0.0
      ValueText = ''
      IsNumeric = $false
    }
  }

  if ($value -notmatch '^[Rr$\d\.\,\s-]+$') {
    return [ordered]@{
      Value = 0.0
      ValueText = $value
      IsNumeric = $false
    }
  }

  $numberText = $value -replace '\s', ''
  $numberText = $numberText -replace '[Rr]\$', ''
  $numberText = $numberText -replace '[^\d,.\-]', ''

  if (-not $numberText) {
    return [ordered]@{
      Value = 0.0
      ValueText = ''
      IsNumeric = $false
    }
  }

  $negative = $numberText.StartsWith('-')
  if ($negative) {
    $numberText = $numberText.Substring(1)
  }

  $lastComma = $numberText.LastIndexOf(',')
  $lastDot = $numberText.LastIndexOf('.')
  $decimalIndex = [Math]::Max($lastComma, $lastDot)

  if ($decimalIndex -ge 0) {
    $integerPart = ($numberText.Substring(0, $decimalIndex) -replace '[^\d]', '')
    $decimalPart = ($numberText.Substring($decimalIndex + 1) -replace '[^\d]', '')

    if (-not $integerPart) {
      $integerPart = '0'
    }

    if (-not $decimalPart) {
      $decimalPart = '00'
    } elseif ($decimalPart.Length -eq 1) {
      $decimalPart += '0'
    } elseif ($decimalPart.Length -gt 2) {
      $decimalPart = $decimalPart.Substring(0, 2)
    }

    $normalized = ('{0}{1}.{2}' -f ($(if ($negative) { '-' } else { '' })), $integerPart, $decimalPart)
  } else {
    $digits = ($numberText -replace '[^\d]', '')
    if (-not $digits) {
      $digits = '0'
    }

    $normalized = ('{0}{1}' -f ($(if ($negative) { '-' } else { '' })), $digits)
  }

  $parsedValue = 0.0
  $ok = [double]::TryParse(
    $normalized,
    [System.Globalization.NumberStyles]::AllowLeadingSign -bor [System.Globalization.NumberStyles]::AllowDecimalPoint,
    [System.Globalization.CultureInfo]::InvariantCulture,
    [ref]$parsedValue
  )

  if (-not $ok) {
    return [ordered]@{
      Value = 0.0
      ValueText = $value
      IsNumeric = $false
    }
  }

  return [ordered]@{
    Value = [double]$parsedValue
    ValueText = ''
    IsNumeric = $true
  }
}

function Resolve-PregaoFornecedorValor {
  param(
    [AllowNull()]
    [string]$ColumnFive,
    [AllowNull()]
    [string]$ColumnSix
  )

  $colFive = Normalize-PlainText $ColumnFive
  $colSix = Normalize-PlainText $ColumnSix
  $valueFive = Convert-ContractValue $colFive
  $valueSix = Convert-ContractValue $colSix

  if ($valueFive.IsNumeric -and -not $valueSix.IsNumeric) {
    return [ordered]@{
      Fornecedor = $(if (Test-LooksLikeStatus $colSix) { '' } else { $colSix })
      ValorInfo = $valueFive
    }
  }

  if ($valueSix.IsNumeric -and -not $valueFive.IsNumeric) {
    return [ordered]@{
      Fornecedor = $(if (Test-LooksLikeStatus $colFive) { '' } else { $colFive })
      ValorInfo = $valueSix
    }
  }

  if ($valueSix.IsNumeric) {
    return [ordered]@{
      Fornecedor = $(if (Test-LooksLikeStatus $colFive) { '' } else { $colFive })
      ValorInfo = $valueSix
    }
  }

  if ($valueFive.IsNumeric) {
    return [ordered]@{
      Fornecedor = $(if (Test-LooksLikeStatus $colSix) { '' } else { $colSix })
      ValorInfo = $valueFive
    }
  }

  return [ordered]@{
    Fornecedor = $(if (Test-LooksLikeStatus $colFive) { '' } elseif ($colFive) { $colFive } elseif (Test-LooksLikeStatus $colSix) { '' } else { $colSix })
    ValorInfo = [ordered]@{
      Value = 0.0
      ValueText = ''
      IsNumeric = $false
    }
  }
}

function Convert-WorksheetRowToContract {
  param(
    [hashtable]$SheetConfig,
    [string[]]$Cells
  )

  $map = $SheetConfig.Mapping
  $numero = Get-TextByIndex -Cells $Cells -Index $map.Numero

  if ($SheetConfig.Mode -eq 'Pregao') {
    $resolved = Resolve-PregaoFornecedorValor -ColumnFive (Get-TextByIndex -Cells $Cells -Index 5) -ColumnSix (Get-TextByIndex -Cells $Cells -Index 6)
    $fornecedor = Normalize-PlainText $resolved.Fornecedor
    $valueInfo = $resolved.ValorInfo
  } else {
    $fornecedor = Get-TextByIndex -Cells $Cells -Index $map.Fornecedor
    $valueInfo = Convert-ContractValue (Get-TextByIndex -Cells $Cells -Index $map.Valor)
  }

  if (Test-LooksLikeStatus $fornecedor) {
    $fornecedor = ''
  }

  return [ordered]@{
    ano = Get-AnoFromNumero $numero
    numero = $numero
    fornecedor = $fornecedor
    objeto = Get-TextByIndex -Cells $Cells -Index $map.Objeto
    processo = Get-TextByIndex -Cells $Cells -Index $map.Processo
    categoria = $SheetConfig.Category
    modalidade = Get-TextByIndex -Cells $Cells -Index $map.Modalidade
    tipo = Get-TipoFromNumero $numero
    valor = [double]$valueInfo.Value
    valor_texto = $valueInfo.ValueText
    inicio_vigencia = Convert-DateToIso (Get-TextByIndex -Cells $Cells -Index $map.Inicio)
    fim_vigencia = Convert-DateToIso (Get-TextByIndex -Cells $Cells -Index $map.Fim)
    status_excel = Get-StatusFromCells $Cells
    observacoes = Get-TextByIndex -Cells $Cells -Index $map.Observacoes
    gestor_fiscal = Get-TextByIndex -Cells $Cells -Index $map.Gestor
  }
}

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$scriptDirectory = Split-Path -Parent $PSCommandPath
if (-not $OutputPath) {
  $OutputPath = Join-Path (Split-Path -Parent $scriptDirectory) 'contratos-data.js'
}
$resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = Split-Path -Parent $resolvedOutputPath

if (-not (Test-Path -LiteralPath $outputDirectory)) {
  [void](New-Item -ItemType Directory -Path $outputDirectory)
}

$workbookItem = Get-Item -LiteralPath $resolvedWorkbookPath
$excel = $null
$workbook = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($resolvedWorkbookPath, $false, $true)

  $contracts = New-Object System.Collections.Generic.List[object]
  $countsByCategory = [ordered]@{}

  for ($sheetIndex = 1; $sheetIndex -le $workbook.Worksheets.Count; $sheetIndex++) {
    $worksheet = $workbook.Worksheets.Item($sheetIndex)
    $usedRange = $null

    try {
      $config = Get-SheetConfig ([string]$worksheet.Name)
      if ($null -eq $config) {
        continue
      }

      $usedRange = $worksheet.UsedRange
      $rowCount = [int]$usedRange.Rows.Count
      $categoryCount = 0

      for ($rowIndex = $config.StartRow; $rowIndex -le $rowCount; $rowIndex++) {
        $cells = Get-RowCells -Worksheet $worksheet -Row $rowIndex -MaxColumns ([int]$config.MaxColumns)
        if (-not (Test-RowHasData $cells)) {
          continue
        }

        $contracts.Add((Convert-WorksheetRowToContract -SheetConfig $config -Cells $cells))
        $categoryCount++
      }

      $countsByCategory[$config.Category] = $categoryCount
    }
    finally {
      if ($usedRange -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange)
      }

      if ($worksheet -ne $null) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet)
      }
    }
  }

  $payload = [ordered]@{
    ultimaAtualizacao = $workbookItem.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:ss')
    origemArquivo = $workbookItem.Name
    contratos = $contracts
  }

  $json = $payload | ConvertTo-Json -Depth 6 -Compress
  $output = @(
    '// Gerado automaticamente por scripts/gerar-dados-contratos.ps1'
    "window.PAINEL_CONTRATOS_DATA = $json;"
    ''
  ) -join [Environment]::NewLine

  [System.IO.File]::WriteAllText($resolvedOutputPath, $output, (New-Object System.Text.UTF8Encoding($false)))

  Write-Output ('Arquivo gerado: ' + $resolvedOutputPath)
  Write-Output ('Origem: ' + $workbookItem.FullName)
  Write-Output ('Total de registros: ' + $contracts.Count)
  foreach ($pair in $countsByCategory.GetEnumerator()) {
    Write-Output (' - ' + $pair.Key + ': ' + $pair.Value)
  }
}
finally {
  if ($workbook -ne $null) {
    $workbook.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
  }

  if ($excel -ne $null) {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
