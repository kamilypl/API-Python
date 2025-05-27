from flask import Flask, request, send_file
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import tempfile
import os

app = Flask(__name__)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

def aplicar_texto_formatado(shape, chave, valor):
    """Aplica estilos diferentes dependendo da chave."""
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = str(valor).replace("\\n", "\n")

    if "titulo" in chave:
        run.font.bold = True
        run.font.size = Pt(20)
    elif "resumo" in chave:
        run.font.size = Pt(14)
    elif "data" in chave:
        run.font.italic = True
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(120, 120, 120)
    elif "link" in chave:
        run.font.size = Pt(12)
        run.font.underline = True
        run.font.color.rgb = RGBColor(0, 102, 204)

@app.route("/gerar_pptx", methods=["POST"])
def gerar_pptx():
    try:
        data = request.json
        print("🔹 Dados recebidos:", data)

        prs = Presentation(TEMPLATE_PATH)

        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for chave, valor in data.items():
                        marcador = f"{{{{{chave}}}}}"
                        if shape.text.strip() == marcador:
                            aplicar_texto_formatado(shape, chave, valor)

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        prs.save(temp_file.name)

        return send_file(
            temp_file.name,
            as_attachment=True,
            download_name="noticias_geradas.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        print("🔥 Erro interno:", str(e))
        return {"erro": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
