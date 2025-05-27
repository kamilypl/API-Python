from flask import Flask, request, send_file
from pptx import Presentation
import tempfile
import os

app = Flask(__name__)
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

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
                        marcador = f"{{{{{chave}}}}}"  # ex: {{titulo0}}
                        if marcador in shape.text:
                            shape.text = shape.text.replace(marcador, valor)

        # Salvar o resultado em um arquivo temporário
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
