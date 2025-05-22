from flask import Flask, request, send_file
from pptx import Presentation
import tempfile
import os

app = Flask(__name__)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

@app.route("/gerar_pptx", methods=["POST"])
def gerar_pptx():
    try:
        payload = request.json
        noticias = payload.get("noticias", [])

        if not noticias:
            return {"erro": "Nenhuma notícia recebida"}, 400

        # Carrega o template existente
        prs = Presentation(TEMPLATE_PATH)
        slide = prs.slides[0]
        preenchidos = 0

        # Preenche os placeholders existentes no slide com os dados das notícias
        for shape in slide.shapes:
            if not shape.has_text_frame or preenchidos >= len(noticias):
                continue

            noticia = noticias[preenchidos]
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.text = (
                        run.text
                        .replace("{{titulo}}", noticia.get("titulo", ""))
                        .replace("{{resumo}}", "\n" + noticia.get("resumo", ""))
                        .replace("{{data}}", "\nData: " + noticia.get("data", ""))
                        .replace("{{link}}", "\n" + noticia.get("link", ""))
                    )

            preenchidos += 1

        print(f"✅ Blocos preenchidos com sucesso: {preenchidos}")

        # Cria arquivo temporário com a apresentação gerada
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        prs.save(temp_file.name)

        return send_file(
            temp_file.name,
            as_attachment=True,
            download_name="noticias_atualizadas.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        print("🔥 Erro ao gerar apresentação:", str(e))
        return {"erro": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
