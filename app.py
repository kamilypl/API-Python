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

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            text_frame = shape.text_frame
            # Concatenar todo o texto da shape
            full_text = ""
            for para in text_frame.paragraphs:
                for run in para.runs:
                    full_text += run.text

            # Substituir todos os placeholders
            for i, noticia in enumerate(noticias):
                full_text = full_text.replace(f"{{{{titulo{i}}}}}", noticia.get("titulo", ""))
                full_text = full_text.replace(f"{{{{resumo{i}}}}}", noticia.get("resumo", ""))
                full_text = full_text.replace(f"{{{{data{i}}}}}", noticia.get("data", ""))
                full_text = full_text.replace(f"{{{{link{i}}}}}", noticia.get("link", ""))

            # Limpar o frame e reinserir o texto substituído
            text_frame.clear()
            p = text_frame.paragraphs[0]
            r = p.add_run()
            r.text = full_text

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
