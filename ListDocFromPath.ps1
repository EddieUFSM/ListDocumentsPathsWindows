# Solicitar ao usuário o nome do arquivo Excel
$excelFileName = Read-Host "Digite o nome do arquivo Excel a ser criado (sem extensão)"

# Verificar se o nome do arquivo foi fornecido
if (-not $excelFileName) {
    Write-Host "Nome do arquivo inválido. Saindo do script."
    exit
}

# Adicionar a extensão .xlsx se não estiver presente
if (-not $excelFileName.EndsWith(".xlsx")) {
    $excelFileName += ".xlsx"
}

# Obter o diretório para varrer a partir do usuário
$diretorioVarredura = Read-Host "Digite o diretório que deseja varrer (por exemplo, 'C:\Downloads')"

# Verificar se o diretório fornecido existe
if (-not (Test-Path $diretorioVarredura)) {
    Write-Host "O diretório fornecido não existe. Saindo do script."
    exit
}

# Criar um arquivo Excel
$excelFilePath = Join-Path -Path $diretorioVarredura -ChildPath $excelFileName
$spreadsheet = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($excelFilePath, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
$workbookPart = $spreadsheet.AddWorkbookPart()
$workbookPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
$worksheetPart = $workbookPart.AddNewPart([DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::GeneratePartUri("Sheet1"))
$worksheetPart.Worksheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet(New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData)

# Adicionar cabeçalhos ao arquivo Excel
$row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
$row.Append((New-Object DocumentFormat.OpenXml.Spreadsheet.Cell -Property @{ DataType = "String"; CellValue = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue "Nome do Arquivo" }))
$row.Append((New-Object DocumentFormat.OpenXml.Spreadsheet.Cell -Property @{ DataType = "String"; CellValue = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue "Caminho" }))
$row.Append((New-Object DocumentFormat.OpenXml.Spreadsheet.Cell -Property @{ DataType = "String"; CellValue = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue "Tamanho (Bytes)" }))
$worksheetPart.Worksheet.SheetData.Append($row)

# Iterar sobre os arquivos no diretório
Get-ChildItem -Path $diretorioVarredura -File -Recurse | ForEach-Object {
    $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
    $row.Append((New-Object DocumentFormat.OpenXml.Spreadsheet.Cell -Property @{ DataType = "String"; CellValue = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue $_.Name }))
    $row.Append((New-Object DocumentFormat.OpenXml.Spreadsheet.Cell -Property @{ DataType = "String"; CellValue = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue $_.FullName }))
    $row.Append((New-Object DocumentFormat.OpenXml.Spreadsheet.Cell -Property @{ DataType = "Number"; CellValue = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue $_.Length }))
    $worksheetPart.Worksheet.SheetData.Append($row)
}

# Salvar e fechar o arquivo Excel
$worksheetPart.Worksheet.Save()
$workbookPart.Workbook.Save()
$spreadsheet.Close()

Write-Host "Arquivo Excel criado com sucesso em: $excelFilePath"
