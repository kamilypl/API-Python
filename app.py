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

        for i, noticia in enumerate(noticias):
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue

                text_frame = shape.text_frame
                text = text_frame.text

                # Substituir apenas se o placeholder estiver contido no texto
                if f"{{{{titulo{i}}}}}" in text:
                    text_frame.text = text.replace(f"{{{{titulo{i}}}}}", noticia.get("titulo", ""))
                if f"{{{{resumo{i}}}}}" in text:
                    text_frame.text = text_frame.text.replace(f"{{{{resumo{i}}}}}", noticia.get("resumo", ""))
                if f"{{{{data{i}}}}}" in text:
                    text_frame.text = text_frame.text.replace(f"{{{{data{i}}}}}", noticia.get("data", ""))
                if f"{{{{link{i}}}}}" in text:
                    text_frame.text = text_frame.text.replace(f"{{{{link{i}}}}}", noticia.get("link", ""))

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
