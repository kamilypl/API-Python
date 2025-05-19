from flask import Flask, request, send_file
from pptx import Presentation
import tempfile
import os

app = Flask(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

@app.route("/gerar_pptx", methods=["POST"])
def gerar_pptx():
    data = request.json
    prs = Presentation(TEMPLATE_PATH)

    # Cria UM SLIDE por entrada enviada (você pode personalizar para vários slides, se enviar lista)
    slide_layout = prs.slide_layouts[1]  # Título + conteúdo

    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = f"{data.get('setor', '')} – {data.get('titulo', '')}"
    corpo = slide.placeholders[1]
    corpo.text = f"{data.get('resumo', '')}\n\n{data.get('link', '')}\nData: {data.get('data', '')}"

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(temp_file.name)

    return send_file(temp_file.name, as_attachment=True, download_name="noticia.pptx",
                     mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
