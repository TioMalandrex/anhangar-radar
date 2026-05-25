import sys

if sys.version_info < (3, 8):
    print(
        f"Erro: Python 3.8 ou superior é necessário.\n"
        f"Versão atual: {sys.version}\n"
        "Baixe em: https://www.python.org/downloads/"
    )
    sys.exit(1)

sys.path.insert(0, "src")
from app import AnhangaRadar

if __name__ == "__main__":
    app = AnhangaRadar()
    app.mainloop()
