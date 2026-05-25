import os
import tempfile
from pathlib import Path
from typing import Optional, Union

from flask import Flask, current_app, jsonify, render_template, request, send_file

import config
from excel_manager import adicionar_linha, contar_registros_hoje, garantir_excel
from ia_processor import extrair_dados
from importador import importar_planilha

ROOT = Path(__file__).parent.parent
DEFAULT_EXCEL_PATH = ROOT / "data" / "contatos.xlsx"

FONTES = ["Ligação", "WhatsApp", "E-mail", "Visita", "Outro"]
STATUSES = [
    "Sem resposta",
    "Não atendida",
    "Número inexistente",
    "Falha no contato",
    "Sem interesse",
    "Contato realizado",
    "Agendado",
]


def _aba_para_status(status: str) -> str:
    s = status.lower()
    if any(k in s for k in ["interesse", "recusou"]):
        return "Sem Interesse"
    return "Falhas e Sem Contato"


def _resolve_api_key(payload=None, form=None) -> str:
    key = ""
    if payload:
        key = (payload.get("api_key") or "").strip()
    if not key and form:
        key = (form.get("api_key") or "").strip()
    if not key:
        key = (os.environ.get("GROQ_API_KEY") or "").strip()
    if not key:
        key = config.carregar_api_key(ROOT)
    return key


def _excel_path() -> Path:
    return Path(current_app.config["EXCEL_PATH"])


def _mensagem_ia_segura(erro: Exception) -> str:
    msg = str(erro).lower()
    if "limite" in msg or "429" in msg:
        return "Limite de requisições atingido. Aguarde alguns segundos e tente novamente."
    if "api" in msg and ("inválida" in msg or "invalid" in msg or "401" in msg):
        return "Chave da API inválida. Verifique sua chave em console.groq.com"
    return "Erro ao chamar a IA. Verifique a API key e tente novamente."


def _mensagem_importacao_segura(erro: Exception) -> str:
    msg = str(erro).lower()
    if "não contém dados" in msg or "vazia" in msg:
        return "A planilha não contém dados válidos."
    return "Erro ao importar planilha."


def create_app(excel_path: Optional[Union[Path, str]] = None) -> Flask:
    app = Flask(
        __name__,
        template_folder=str(ROOT / "web" / "templates"),
        static_folder=str(ROOT / "web" / "static"),
    )
    app.config["EXCEL_PATH"] = str(excel_path or DEFAULT_EXCEL_PATH)

    @app.get("/")
    def index():
        return render_template("index.html", fontes=FONTES, statuses=STATUSES)

    @app.get("/api/contador")
    def contador():
        total = contar_registros_hoje(str(_excel_path()))
        return jsonify(total=total)

    @app.post("/api/processar")
    def processar():
        payload = request.get_json(silent=True) or {}
        texto = (payload.get("texto") or "").strip()
        if not texto:
            return jsonify(error="Descreva o contato antes de processar."), 400
        api_key = _resolve_api_key(payload=payload)
        if not api_key:
            return jsonify(error="API key não informada."), 400
        try:
            dados = extrair_dados(texto, api_key)
        except ValueError as exc:
            return jsonify(error=_mensagem_ia_segura(exc)), 400
        except Exception:
            current_app.logger.exception("Erro ao processar texto via IA.")
            return jsonify(error="Erro ao chamar a IA."), 500
        return jsonify(dados=dados)

    @app.post("/api/salvar")
    def salvar():
        payload = request.get_json(silent=True) or {}
        nome = (payload.get("nome") or "").strip()
        if not nome:
            return jsonify(error="O campo Nome está vazio."), 400

        dados = {
            "nome": nome,
            "sobrenome": (payload.get("sobrenome") or "").strip(),
            "cidade": (payload.get("cidade") or "").strip(),
            "telefone": (payload.get("telefone") or "").strip(),
            "fonte": (payload.get("fonte") or "").strip(),
            "status": (payload.get("status") or "").strip(),
            "observacoes": (payload.get("observacoes") or "").strip(),
        }
        dados["aba"] = _aba_para_status(dados["status"])

        try:
            adicionar_linha(str(_excel_path()), dados)
        except PermissionError as exc:
            return jsonify(
                error="O arquivo 'contatos.xlsx' está aberto em outro programa. "
                "Feche o Excel e tente novamente."
            ), 409
        except Exception:
            current_app.logger.exception("Erro ao salvar registro.")
            return jsonify(error="Erro ao salvar o registro."), 500

        contador_atual = contar_registros_hoje(str(_excel_path()))
        nome_completo = f'{dados["nome"]} {dados["sobrenome"]}'.strip()
        return jsonify(
            message=f'✓ "{nome_completo}" salvo em "{dados["aba"]}".',
            contador=contador_atual,
        )

    @app.post("/api/importar")
    def importar():
        arquivo = request.files.get("arquivo")
        if not arquivo:
            return jsonify(error="Selecione um arquivo .xlsx."), 400
        api_key = _resolve_api_key(form=request.form)
        if not api_key:
            return jsonify(error="API key não informada."), 400

        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                arquivo.save(tmp.name)
                temp_path = tmp.name

            importados, erros = importar_planilha(
                temp_path, str(_excel_path()), api_key
            )
        except ValueError as exc:
            return jsonify(error=_mensagem_importacao_segura(exc)), 400
        except Exception:
            current_app.logger.exception("Erro ao importar planilha.")
            return jsonify(error="Erro ao importar planilha."), 500
        finally:
            if temp_path and os.path.exists(temp_path):
                os.unlink(temp_path)

        contador_atual = contar_registros_hoje(str(_excel_path()))
        return jsonify(
            importados=importados,
            contador=contador_atual,
        )

    @app.get("/api/baixar")
    def baixar():
        garantir_excel(str(_excel_path()))
        return send_file(
            str(_excel_path()),
            as_attachment=True,
            download_name="contatos.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    return app
