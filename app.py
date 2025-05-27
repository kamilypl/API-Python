from flask import Flask, request, send_file
from pptx import Presentation
import io

app = Flask(__name__)

@app.route('/gerar_ppt', methods=['POST'])
def gerar_ppt():
    dados = request.json

    # Carrega o template
    prs = Presentation('template.pptx')

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text
                for chave, valor in dados.items():
                    if f"{{{{{chave}}}}}" in text:
                        shape.text = text.replace(f"{{{{{chave}}}}}", valor)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    return send_file(output, download_name="output.pptx", as_attachment=True)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
