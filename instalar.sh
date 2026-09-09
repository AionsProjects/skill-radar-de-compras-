#!/usr/bin/env bash
# Instalador do Radar de Compras (skill busca-preco) para macOS e Linux.
#
# Resolve o caso de quem NAO tem Python: encontra um interpretador que sirva,
# diz como instalar se nao houver, instala as dependencias e roda a verificacao.
#
# Uso:  bash instalar.sh
set -u

MINIMO_MAIOR=3
MINIMO_MENOR=10
PASTA="$HOME/.claude/skills/busca-preco"
DEPS=(requests beautifulsoup4 lxml openpyxl reportlab)

titulo() { printf '\n\033[36m%s\033[0m\n' "$1"; }
ok()     { printf '  \033[32mok   \033[0m %s\n' "$1"; }
aviso()  { printf '  \033[33maviso\033[0m %s\n' "$1"; }
erro()   { printf '  \033[31mERRO \033[0m %s\n' "$1"; }

# ---------------------------------------------------------------------------
# 1. Achar um Python que sirva
# ---------------------------------------------------------------------------
# Valida rodando `--version` de verdade: `command -v` encontrar o binario nao
# garante que ele execute nem que a versao sirva.
PY=""
achar_python() {
  for cand in python3 python python3.13 python3.12 python3.11 python3.10; do
    command -v "$cand" >/dev/null 2>&1 || continue
    v="$("$cand" -c 'import sys; print("%d %d" % sys.version_info[:2])' 2>/dev/null)" || continue
    [ -n "$v" ] || continue
    maior="${v% *}"; menor="${v#* }"
    if [ "$maior" -gt "$MINIMO_MAIOR" ] 2>/dev/null || \
       { [ "$maior" -eq "$MINIMO_MAIOR" ] && [ "$menor" -ge "$MINIMO_MENOR" ]; } 2>/dev/null; then
      PY="$cand"
      PY_VERSAO="$maior.$menor"
      return 0
    fi
    aviso "$cand e Python $maior.$menor -- a skill precisa de $MINIMO_MAIOR.$MINIMO_MENOR ou mais novo"
  done
  return 1
}

printf '\033[1mRadar de Compras -- instalador\033[0m\n'
titulo "Procurando Python $MINIMO_MAIOR.$MINIMO_MENOR ou mais novo"

if ! achar_python; then
  erro 'nenhum Python utilizavel encontrado'
  echo
  if [ "$(uname -s)" = "Darwin" ]; then
    echo '  No macOS, o mais simples:'
    if command -v brew >/dev/null 2>&1; then
      echo '    brew install python@3.12'
    else
      echo '    1) instale o Homebrew:  https://brew.sh'
      echo '    2) brew install python@3.12'
      echo '  Ou baixe de https://www.python.org/downloads/macos/'
    fi
  else
    echo '  No Linux, conforme a distribuicao:'
    echo '    Ubuntu/Debian:  sudo apt install python3 python3-pip'
    echo '    Fedora:         sudo dnf install python3 python3-pip'
    echo '    Arch:           sudo pacman -S python python-pip'
  fi
  echo
  echo '  Depois rode este instalador de novo.'
  exit 1
fi
ok "Python $PY_VERSAO via '$PY'"

# ---------------------------------------------------------------------------
# 2. Dependencias
# ---------------------------------------------------------------------------
titulo 'Instalando as dependencias'
echo "  ${DEPS[*]}"
if ! "$PY" -m pip --version >/dev/null 2>&1; then
  erro 'este Python nao tem pip'
  echo '    Ubuntu/Debian: sudo apt install python3-pip'
  echo '    macOS:         python3 -m ensurepip --upgrade'
  exit 1
fi

# Distribuicoes novas (PEP 668) recusam instalar no Python do sistema. O
# --user resolve sem exigir sudo nem quebrar pacotes da distro.
if ! "$PY" -m pip install --quiet --user "${DEPS[@]}" 2>/dev/null; then
  aviso 'pip --user falhou; tentando com --break-system-packages'
  if ! "$PY" -m pip install --quiet --user --break-system-packages "${DEPS[@]}"; then
    erro 'pip falhou. Rode sem --quiet para ver o motivo:'
    echo "    $PY -m pip install --user ${DEPS[*]}"
    echo '  Se a distribuicao bloquear, um ambiente virtual resolve:'
    echo "    $PY -m venv ~/.venv-busca-preco"
    echo "    ~/.venv-busca-preco/bin/pip install ${DEPS[*]}"
    echo "    ~/.venv-busca-preco/bin/python $PASTA/busca_preco.py --doctor"
    exit 1
  fi
fi
ok 'dependencias instaladas'

# ---------------------------------------------------------------------------
# 3. Verificacao final
# ---------------------------------------------------------------------------
SCRIPT="$PASTA/busca_preco.py"
if [ ! -f "$SCRIPT" ]; then
  AQUI="$(cd "$(dirname "$0")" && pwd)"
  if [ -f "$AQUI/busca_preco.py" ]; then
    SCRIPT="$AQUI/busca_preco.py"
    aviso "usando $AQUI (a skill so aparece no Claude Code se estiver em $PASTA)"
  else
    erro "nao achei busca_preco.py em $PASTA"
    echo
    echo '  Clone a skill primeiro, com este destino exato:'
    echo "    git clone https://github.com/AionsProjects/skill-radar-de-compras-.git \"$PASTA\""
    exit 1
  fi
fi

titulo 'Verificando o ambiente'
"$PY" "$SCRIPT" --doctor
codigo=$?

echo
if [ $codigo -eq 0 ]; then
  printf '\033[32mTudo pronto.\033[0m\n\n'
  echo '  1. Reinicie o Claude Code (a skill so aparece em sessao nova).'
  echo '  2. Use:  /busca-preco minha-lista.xlsx AM'
  echo
  echo "  Planilha de exemplo: $PASTA/referencias/exemplo-amazonas.xlsx"
  echo '  O Amazonas funciona agora. A Paraiba precisa da extensao do Claude'
  echo '  no Chrome (https://claude.ai/chrome) -- veja o README.'
else
  erro 'a verificacao apontou problemas -- veja acima o que falta'
fi
exit $codigo
