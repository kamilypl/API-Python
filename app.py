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
        print("🧾 JSON recebido:")
        print(dados)

        noticias = dados.get("noticias", [])

        prs = Presentation(TEMPLATE_PATH)
        slide = prs.slides[0]

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for i, noticia in enumerate(noticias):
                        run.text = run.text.replace(f"{{{{titulo{i}}}}}", noticia.get("titulo", ""))
                        run.text = run.text.replace(f"{{{{resumo{i}}}}}", noticia.get("resumo", ""))
                        run.text = run.text.replace(f"{{{{data{i}}}}}", noticia.get("data", ""))
                        run.text = run.text.replace(f"{{{{link{i}}}}}", noticia.get("link", ""))

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
