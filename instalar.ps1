# Instalador do Radar de Compras (skill busca-preco) para Windows.
#
# Resolve o caso de quem NAO tem Python: encontra um interpretador que sirva,
# oferece instalar um se nao houver, instala as dependencias e roda a
# verificacao final.
#
# Uso:  powershell -ExecutionPolicy Bypass -File instalar.ps1
#
# Nao exige administrador. Nao mexe em nada fora da pasta da skill e do proprio
# Python do usuario.

$ErrorActionPreference = 'Stop'
$MINIMO = [version]'3.10'
$PASTA  = Join-Path $env:USERPROFILE '.claude\skills\busca-preco'
$DEPS   = @('requests', 'beautifulsoup4', 'lxml', 'openpyxl', 'reportlab')

function Titulo($t) { Write-Host ''; Write-Host $t -ForegroundColor Cyan }
function Ok($t)     { Write-Host "  ok    $t" -ForegroundColor Green }
function Aviso($t)  { Write-Host "  aviso $t" -ForegroundColor Yellow }
function Erro($t)   { Write-Host "  ERRO  $t" -ForegroundColor Red }

# ---------------------------------------------------------------------------
# 1. Achar um Python que REALMENTE funcione
# ---------------------------------------------------------------------------
# Cuidado com o stub da Microsoft Store: no Windows 10/11 existe
# WindowsApps\python3.exe, que esta no PATH, responde a Get-Command e ao ser
# chamado ABRE A LOJA em vez de executar. Por isso cada candidato e validado
# rodando `--version` de verdade e conferindo se saiu um numero.
function Achar-Python {
    $candidatos = @(
        @{ exe = 'py';      args = @('-3') },
        @{ exe = 'python';  args = @() },
        @{ exe = 'python3'; args = @() }
    )
    foreach ($c in $candidatos) {
        $g = Get-Command $c.exe -ErrorAction SilentlyContinue
        if (-not $g) { continue }
        if ($g.Source -like '*WindowsApps*') {
            Aviso "$($c.exe) e o atalho da Microsoft Store, nao um Python real -- ignorado"
            continue
        }
        try {
            $saida = & $c.exe @($c.args + '--version') 2>&1 | Out-String
        } catch { continue }
        if ($saida -match 'Python (\d+)\.(\d+)\.?(\d*)') {
            $v = [version]("$($Matches[1]).$($Matches[2])")
            if ($v -ge $MINIMO) {
                # caminho ABSOLUTO de proposito: com o nome ('py'), o operador &
                # resolveria a funcao homonima deste script antes do executavel,
                # e o instalador entrava em recursao infinita
                return @{ exe = $g.Source; nome = $c.exe; args = $c.args; versao = $v }
            }
            Aviso "$($c.exe) e Python $v -- a skill precisa de $MINIMO ou mais novo"
        }
    }
    return $null
}

Write-Host 'Radar de Compras -- instalador' -ForegroundColor White
Titulo 'Procurando Python 3.10 ou mais novo'
$py = Achar-Python

if (-not $py) {
    Erro 'nenhum Python utilizavel encontrado'
    $temWinget = [bool](Get-Command winget -ErrorAction SilentlyContinue)
    if ($temWinget) {
        Write-Host ''
        Write-Host '  Posso instalar o Python 3.12 agora, pela loja oficial da Microsoft'
        Write-Host '  (winget install Python.Python.3.12). Leva 1 a 2 minutos.'
        $r = Read-Host '  Instalar? (s/n)'
        if ($r -notmatch '^[sSyY]') {
            Write-Host ''
            Write-Host '  Sem problema. Instale manualmente em https://www.python.org/downloads/'
            Write-Host '  e MARQUE "Add python.exe to PATH" na primeira tela. Depois rode este'
            Write-Host '  instalador de novo.'
            exit 1
        }
        Titulo 'Instalando Python 3.12'
        winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements
        # o PATH da sessao atual nao ve o que acabou de ser instalado
        $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                    [Environment]::GetEnvironmentVariable('Path', 'User')
        $py = Achar-Python
        if (-not $py) {
            Erro 'instalei, mas o Python ainda nao aparece nesta janela'
            Write-Host ''
            Write-Host '  Isso e normal: o PATH so atualiza em janelas novas.'
            Write-Host '  FECHE este PowerShell, abra outro e rode este instalador de novo.'
            exit 1
        }
    } else {
        Write-Host ''
        Write-Host '  Este Windows nao tem winget. Instale o Python manualmente:'
        Write-Host '    https://www.python.org/downloads/'
        Write-Host '  MARQUE "Add python.exe to PATH" na primeira tela do instalador.'
        Write-Host '  Depois rode este instalador de novo.'
        exit 1
    }
}
Ok "Python $($py.versao) via '$($py.nome) $($py.args -join ' ')' -- $($py.exe)"

# NAO chame esta funcao de `Py`: o executavel tambem se chama py, o PowerShell
# resolve funcao antes de programa, e `& $py.exe` viraria recursao infinita.
function Invoke-Python { param([string[]]$Argumentos)
    & $py.exe @($py.args + $Argumentos)
}

# ---------------------------------------------------------------------------
# 2. Dependencias
# ---------------------------------------------------------------------------
Titulo 'Instalando as dependencias'
Write-Host "  $($DEPS -join ', ')"
Invoke-Python @('-m', 'pip', 'install', '--quiet', '--upgrade', 'pip')
Invoke-Python (@('-m', 'pip', 'install', '--quiet') + $DEPS)
if ($LASTEXITCODE -ne 0) {
    Erro 'pip falhou. Rode sem --quiet para ver o motivo:'
    Write-Host "    $($py.nome) $($py.args -join ' ') -m pip install $($DEPS -join ' ')"
    exit 1
}
Ok 'dependencias instaladas'

# ---------------------------------------------------------------------------
# 3. Verificacao final
# ---------------------------------------------------------------------------
$script = Join-Path $PASTA 'busca_preco.py'
if (-not (Test-Path $script)) {
    # o instalador tambem funciona rodado de dentro da pasta clonada
    $local = Join-Path $PSScriptRoot 'busca_preco.py'
    if (Test-Path $local) {
        $script = $local
        Aviso "usando $PSScriptRoot (a skill so aparece no Claude Code se estiver em $PASTA)"
    } else {
        Erro "nao achei busca_preco.py em $PASTA"
        Write-Host ''
        Write-Host '  Clone a skill primeiro, com este destino exato:'
        Write-Host '    git clone https://github.com/AionsProjects/skill-radar-de-compras-.git "' -NoNewline
        Write-Host "$PASTA`""
        exit 1
    }
}

Titulo 'Verificando o ambiente'
Invoke-Python @($script, '--doctor')
$codigo = $LASTEXITCODE

Write-Host ''
if ($codigo -eq 0) {
    Write-Host 'Tudo pronto.' -ForegroundColor Green
    Write-Host ''
    Write-Host '  1. Reinicie o Claude Code (a skill so aparece em sessao nova).'
    Write-Host '  2. Use:  /busca-preco minha-lista.xlsx AM'
    Write-Host ''
    Write-Host "  Planilha de exemplo: $PASTA\referencias\exemplo-amazonas.xlsx"
    Write-Host '  O Amazonas funciona agora. A Paraiba precisa da extensao do Claude'
    Write-Host '  no Chrome (https://claude.ai/chrome) -- veja o README.'
} else {
    Erro 'a verificacao apontou problemas -- veja acima o que falta'
}
exit $codigo
