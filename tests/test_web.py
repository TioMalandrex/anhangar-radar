import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
from web_app import create_app


class TestWebRoutes(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.excel_path = os.path.join(self.temp_dir.name, "contatos.xlsx")
        self.app = create_app(excel_path=self.excel_path)
        self.client = self.app.test_client()

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_contador_inicial(self):
        response = self.client.get("/api/contador")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["total"], 0)

    def test_salvar_e_contador(self):
        payload = {
            "nome": "Maria",
            "sobrenome": "Silva",
            "cidade": "Campinas",
            "telefone": "11999990000",
            "fonte": "Ligação",
            "status": "Sem interesse",
            "observacoes": "Teste",
        }
        response = self.client.post("/api/salvar", json=payload)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["contador"], 1)

        response = self.client.get("/api/contador")
        self.assertEqual(response.get_json()["total"], 1)

    def test_baixar_excel(self):
        response = self.client.get("/api/baixar")
        self.assertEqual(response.status_code, 200)
        self.assertIn("contatos.xlsx", response.headers.get("Content-Disposition", ""))
        response.close()

    def test_salvar_sem_nome_retorna_erro(self):
        response = self.client.post("/api/salvar", json={"nome": ""})
        self.assertEqual(response.status_code, 400)


if __name__ == "__main__":
    unittest.main(verbosity=2)
