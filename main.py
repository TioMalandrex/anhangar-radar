import os
import sys
from pathlib import Path

if sys.version_info < (3, 8):
    print(
        f"Erro: Python 3.8 ou superior é necessário.\n"
        f"Versão atual: {sys.version}\n"
        "Baixe em: https://www.python.org/downloads/"
    )
    sys.exit(1)

sys.path.insert(0, str(Path(__file__).parent / "src"))
from web_app import create_app

if __name__ == "__main__":
    app = create_app()
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    debug = os.environ.get("FLASK_DEBUG") == "1"
    app.run(host=host, port=port, debug=debug)
