# Solicitar ao usuario o diretorio para salvar o arquivo CSV
$diretorioSalvamento = Read-Host "Digite o diretorio onde deseja salvar o arquivo CSV"

# Verificar se o diretorio de salvamento fornecido existe
if (-not (Test-Path $diretorioSalvamento)) {
    Write-Host "O diretorio de salvamento fornecido nao existe. Saindo do script."
    exit
}

# Solicitar ao usuario o diretorio a ser varrido (por exemplo, 'E:\')
$diretorioVarredura = Read-Host "Digite o diretorio que deseja varrer (por exemplo, 'E:\')"

# Verificar se o diretorio fornecido existe
if (-not (Test-Path $diretorioVarredura)) {
    Write-Host "O diretorio fornecido nao existe. Saindo do script."
    exit
}

# Solicitar ao usuario a extensao dos arquivos a serem buscados (por exemplo, 'docx')
$extensaoArquivos = Read-Host "Digite a extensao dos arquivos a serem buscados (por exemplo, 'docx')"

# Verificar se o usuario forneceu uma extensao valida
if (-not $extensaoArquivos) {
    Write-Host "Extensao de arquivo invalida. Saindo do script."
    exit
}

# Lista para armazenar os detalhes dos arquivos
$arquivos = @()

# Iterar sobre os arquivos e pastas no diretorio com a extensao especificada
Get-ChildItem -Path $diretorioVarredura -File -Recurse | Where-Object { $_.Extension -eq ".$extensaoArquivos" } | ForEach-Object {
    $arquivo = [PSCustomObject]@{
        Nome = $_.Name
        Caminho = $_.FullName
        TamanhoBytes = $_.Length
    }
    $arquivos += $arquivo
}

# Iterar sobre as pastas no diretorio
Get-ChildItem -Path $diretorioVarredura -Directory -Recurse | ForEach-Object {
    # Iterar sobre os arquivos na pasta atual com a extensao especificada
    Get-ChildItem -Path $_.FullName -File | Where-Object { $_.Extension -eq ".$extensaoArquivos" } | ForEach-Object {
        $arquivo = [PSCustomObject]@{
            Nome = $_.Name
            Caminho = $_.FullName
            TamanhoBytes = $_.Length
        }
        $arquivos += $arquivo
    }
}

# Criar um arquivo CSV com os detalhes dos arquivos
$csvFilePath = Join-Path -Path $diretorioSalvamento -ChildPath "DetalhesArquivos.csv"
$arquivos | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Host "Arquivo CSV criado com sucesso em: $csvFilePath"
