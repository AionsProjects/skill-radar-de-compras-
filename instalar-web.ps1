# Instalador do Radar de Compras que baixa sozinho -- para quem nao tem Git.
#
# Uso (cole no PowerShell, uma linha so):
#
#   irm https://raw.githubusercontent.com/AionsProjects/skill-radar-de-compras-/main/instalar-web.ps1 | iex
#
# Baixa a versao atual, descompacta num temporario e chama instalar.ps1, que
# cuida do Python, das dependencias, da copia para .claude\skills e do registro
# no Claude Desktop.
#
# Serve tambem para ATUALIZAR: rodar de novo troca pela versao mais recente.

$ErrorActionPreference = 'Stop'

$REPO = 'AionsProjects/skill-radar-de-compras-'
$ZIP  = "https://github.com/$REPO/archive/refs/heads/main.zip"

function Titulo($t) { Write-Host ''; Write-Host $t -ForegroundColor Cyan }
function Ok($t)     { Write-Host "  ok    $t" -ForegroundColor Green }
function Erro($t)   { Write-Host "  ERRO  $t" -ForegroundColor Red }

Titulo 'Baixando o Radar de Compras'

$tmp     = Join-Path $env:TEMP ('radar-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$arquivo = Join-Path $tmp 'radar.zip'
New-Item -ItemType Directory -Force -Path $tmp | Out-Null

try {
    # TLS 1.2 explicito: o Windows 10 mais antigo ainda negocia 1.0 por padrao
    # e o GitHub recusa, com um erro que nao explica nada.
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    # a barra de progresso do Invoke-WebRequest deixa o download varias vezes
    # mais lento em PowerShell 5.1
    $progresso = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $ZIP -OutFile $arquivo -UseBasicParsing
    $ProgressPreference = $progresso
    Ok ('baixado ({0:N0} KB)' -f ((Get-Item $arquivo).Length / 1KB))
} catch {
    Erro "nao consegui baixar: $($_.Exception.Message)"
    Write-Host ''
    Write-Host '  Baixe manualmente por aqui:'
    Write-Host "    https://github.com/$REPO"
    Write-Host '  Botao verde Code -> Download ZIP, descompacte e rode instalar.ps1'
    exit 1
}

Titulo 'Descompactando'
Expand-Archive -Path $arquivo -DestinationPath $tmp -Force

$instalador = Get-ChildItem -Path $tmp -Recurse -Filter 'instalar.ps1' |
              Select-Object -First 1
if (-not $instalador) {
    Erro 'o pacote baixado nao tem instalar.ps1'
    exit 1
}
Ok "pacote pronto em $($instalador.Directory.FullName)"

# O instalador real assume que $PSScriptRoot e a pasta da skill, e e: ele esta
# dentro dela. Chamado por caminho completo, isso continua valendo.
& powershell -ExecutionPolicy Bypass -File $instalador.FullName
$codigo = $LASTEXITCODE

# Nao apaga o temporario quando deu errado: e o unico lugar onde da para olhar
# o que veio, se precisar depurar.
if ($codigo -eq 0) {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
} else {
    Write-Host ''
    Write-Host "  Os arquivos baixados ficaram em: $tmp"
}
exit $codigo
