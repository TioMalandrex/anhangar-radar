from groq import Groq
import json


def extrair_dados(texto: str, api_key: str) -> dict:
    client = Groq(api_key=api_key)

    prompt = f"""Extraia informações de contato com vereador e retorne SOMENTE JSON válido, sem markdown.

Campos: nome, sobrenome, cidade, telefone, fonte (Ligação/WhatsApp/E-mail/Instagram/Indicação),
status (use exatamente: Sem resposta, Não atendida, Número inexistente, Falha no contato, Sem interesse, Contato realizado, Agendado),
observacoes. Campos não mencionados retornam string vazia.

Texto: {texto}"""

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
        )
    except Exception as exc:
        msg = str(exc)
        if "rate_limit" in msg.lower() or "429" in msg:
            raise ValueError(
                "Limite de requisições atingido. Aguarde alguns segundos e tente novamente.\n"
                "Limite gratuito: 30 req/min e 14.400 req/dia."
            )
        if "invalid_api_key" in msg.lower() or "401" in msg:
            raise ValueError(
                "Chave da API inválida. Verifique sua chave em console.groq.com"
            )
        raise

    raw = response.choices[0].message.content.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        raise ValueError(f"A IA não retornou JSON válido. Resposta recebida:\n{raw[:300]}")
