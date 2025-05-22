from pptx import Presentation

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

        print(f"🔹 Recebido {len(noticias)} notícias")

        prs = Presentation(TEMPLATE_PATH)
        slide = prs.slides[0]  # Assume que o primeiro slide tem os 6 blocos de texto

        # Preenche os shapes com os dados das notícias, respeitando a ordem
        shape_index = 0
        for noticia in noticias:
            if shape_index >= len(slide.shapes):
                print("⚠️ Mais notícias que caixas de texto disponíveis")
                break

            shape = slide.shapes[shape_index]
            if shape.has_text_frame:
                shape.text_frame.clear()
                shape.text_frame.text = (
                    f"Título: {noticia.get('titulo', '')}\n"
                    f"Data: {noticia.get('data', '')}\n"
                    f"Resumo: {noticia.get('resumo', '')}\n"
                    f"Fonte: {noticia.get('link', '')}"
                )
                shape_index += 1

        # Salva em arquivo temporário
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
