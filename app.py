from flask import Flask, request, send_file
from pptx import Presentation
import tempfile
import os

# Cria a aplicação Flask
app = Flask(__name__)

# Caminho do template .pptx
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

@app.route("/gerar_pptx", methods=["POST"])
def gerar_pptx():
    try:
        # Lê os dados JSON enviados no corpo da requisição
        dados = request.json
        noticias = dados.get("noticias", [])

        # Carrega o arquivo PowerPoint modelo
        prs = Presentation(TEMPLATE_PATH)
        slide = prs.slides[0]  # Considera apenas o primeiro slide

        preenchidos = 0  # Contador de blocos preenchidos

        for shape in slide.shapes:
            if not shape.has_text_frame or preenchidos >= len(noticias):
                continue

            noticia = noticias[preenchidos]

            # Captura o texto atual da caixa
            texto_original = shape.text_frame.text

            # Substitui os placeholders pelos valores da notícia
            texto_formatado = (
                texto_original
                .replace("{{titulo}}", noticia.get("titulo", ""))
                .replace("{{resumo}}", "\n" + noticia.get("resumo", ""))
                .replace("{{data}}", "\nData: " + noticia.get("data", ""))
                .replace("{{link}}", "\n" + noticia.get("link", ""))
            )

            # Limpa a caixa de texto inteira
            shape.text_frame.clear()

            # Insere o texto formatado no primeiro parágrafo
            shape.text_frame.paragraphs[0].text = texto_formatado

            preenchidos += 1

        print(f"✅ Blocos preenchidos: {preenchidos}")

        # Cria um arquivo temporário para salvar a nova apresentação
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        prs.save(temp_file.name)

        # Retorna o arquivo como resposta
        return send_file(
            temp_file.name,
            as_attachment=True,
            download_name="noticias_atualizadas.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        print("🔥 Erro interno:", str(e))
        return {"erro": str(e)}, 500

# Roda o servidor
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
