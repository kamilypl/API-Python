from flask import Flask, request, send_file
from pptx import Presentation
import tempfile
import os

app = Flask(__name__)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

@app.route("/gerar_pptx", methods=["POST"])
def gerar_pptx():
    try:
        dados = request.json
        noticias = dados.get("noticias", [])

        prs = Presentation(TEMPLATE_PATH)
        slide = prs.slides[0]
        preenchidos = 0

        for shape in slide.shapes:
            if not shape.has_text_frame or preenchidos >= len(noticias):
                continue

            noticia = noticias[preenchidos]

            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    texto = run.text
                    texto = texto.replace("{{titulo}}", noticia.get("titulo", ""))
                    texto = texto.replace("{{resumo}}", "\n" + noticia.get("resumo", ""))
                    texto = texto.replace("{{data}}", "\nData: " + noticia.get("data", ""))
                    texto = texto.replace("{{link}}", "\n" + noticia.get("link", ""))
                    run.text = texto

            preenchidos += 1

        print(f"✅ Blocos preenchidos: {preenchidos}")

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        prs.save(temp_file.name)

        return send_file(
            temp_file.name,
            as_attachment=True,
            download_name="noticias_atualizadas.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        print("🔥 Erro interno:", str(e))
        return {"erro": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
