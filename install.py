#!/usr/bin/env python3
import os
import subprocess
import sys
import venv

VENV_DIR    = "venv"
REQUIREMENTS = "requirements.txt"
MIN_PYTHON  = (3, 9)


def _run(cmd: list, msg: str) -> None:
    print(f"  -> {' '.join(cmd)}")
    if subprocess.run(cmd).returncode != 0:
        print(f"\n[ERRO] {msg}")
        sys.exit(1)


def _pip() -> str:
    sub = "Scripts" if os.name == "nt" else "bin"
    exe = "pip.exe" if os.name == "nt" else "pip"
    return os.path.join(VENV_DIR, sub, exe)


def _python() -> str:
    sub = "Scripts" if os.name == "nt" else "bin"
    exe = "python.exe" if os.name == "nt" else "python"
    return os.path.join(VENV_DIR, sub, exe)


def _activate() -> str:
    return (
        rf"{VENV_DIR}\Scripts\activate" if os.name == "nt"
        else f"source {VENV_DIR}/bin/activate"
    )


def main() -> None:
    print("=" * 55)
    print("  Instalador — Plataforma ISSE")
    print("=" * 55)

    if sys.version_info < MIN_PYTHON:
        ver = ".".join(map(str, MIN_PYTHON))
        print(f"\n[ERRO] Python {ver}+ necessário (detectado: {sys.version.split()[0]}).")
        print("       https://www.python.org/downloads/")
        sys.exit(1)

    print(f"\nPython {sys.version.split()[0]} — OK")

    if not os.path.isdir(VENV_DIR):
        print(f"\n[1/3] Criando ambiente virtual em ./{VENV_DIR} ...")
        try:
            venv.create(VENV_DIR, with_pip=True)
        except Exception as exc:
            print(f"\n[ERRO] Não foi possível criar o ambiente virtual: {exc}")
            sys.exit(1)
    else:
        print(f"\n[1/3] Ambiente virtual já existe em ./{VENV_DIR} (ok)")

    pip = _pip()

    if not os.path.isfile(pip):
        print(f"\n[ERRO] pip não encontrado em {pip}.")
        print(f"       Apague a pasta '{VENV_DIR}' e rode o instalador novamente.")
        sys.exit(1)

    print("\n[2/3] Atualizando pip ...")
    _run([pip, "install", "--upgrade", "pip", "-q"], "Falha ao atualizar o pip.")

    if not os.path.isfile(REQUIREMENTS):
        print(f"\n[ERRO] {REQUIREMENTS} não encontrado.")
        print("       Rode o instalador a partir da pasta raiz do projeto.")
        sys.exit(1)

    print(f"\n[3/3] Instalando dependências ...")
    _run(
        [pip, "install", "-r", REQUIREMENTS, "-q"],
        "Falha ao instalar as dependências. Verifique sua conexão e tente novamente.",
    )

    print("\n" + "=" * 55)
    print("  Instalação concluída!")
    print("=" * 55)
    print(f"""
Próximos passos:
  1. {_activate()}
  2. {_python()} backend/main.py
  3. Abra frontend/index.html no navegador.

  Atalho: rode run.bat (Windows) ou ./run.sh (Linux/Mac).
""")


if __name__ == "__main__":
    main()
