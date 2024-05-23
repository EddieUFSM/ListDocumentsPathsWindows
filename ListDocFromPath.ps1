# Solicitar ao usuário o diretório a ser varrido
$diretorioVarredura = Read-Host "Digite o diretório que deseja varrer (por exemplo, 'E:\')"

# Verificar se o diretório fornecido existe
if (-not (Test-Path $diretorioVarredura)) {
    Write-Host "O diretório fornecido não existe. Saindo do script."
    exit
}

# Solicitar ao usuário a extensão dos arquivos a serem buscados
$extensaoArquivos = Read-Host "Digite a extensão dos arquivos a serem buscados (por exemplo, 'docx')"

# Verificar se o usuário forneceu uma extensão válida
if (-not $extensaoArquivos) {
    Write-Host "Extensão de arquivo inválida. Saindo do script."
    exit
}

# Lista para armazenar os detalhes dos arquivos
$arquivos = @()

# Iterar sobre os arquivos no diretório com a extensão especificada
Get-ChildItem -Path $diretorioVarredura -File -Recurse | Where-Object { $_.Extension -eq ".$extensaoArquivos" } | ForEach-Object {
    $arquivo = [PSCustomObject]@{
        Nome = $_.Name
        Caminho = $_.FullName
        TamanhoBytes = $_.Length
    }
    $arquivos += $arquivo
}

# Criar um arquivo CSV com os detalhes dos arquivos
$csvFilePath = Join-Path -Path $diretorioVarredura -ChildPath "DetalhesArquivos.csv"
$arquivos | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Host "Arquivo CSV criado com sucesso em: $csvFilePath"
