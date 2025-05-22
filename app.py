from flask import Flask, request, send_file
from pptx import Presentation
import tempfile
import os

@app.route("/gerar_pptx", methods=["POST"])
def gerar_pptx():
    try:
        req = request.json
        print("🔹 Dados recebidos:", req)

        noticias = req.get("noticias", [])
        if not noticias:
            return {"erro": "Nenhuma notícia recebida"}, 400

        noticia = noticias[0]  # Pega apenas a primeira para o slide atual

        prs = Presentation(TEMPLATE_PATH)
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Slide vazio

        # Substituição manual dos placeholders
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            shape.text = shape.text.replace("{{titulo}}", noticia.get("titulo", ""))
            shape.text = shape.text.replace("{{resumo}}", noticia.get("resumo", ""))
            shape.text = shape.text.replace("{{data}}", noticia.get("data", ""))
            shape.text = shape.text.replace("{{link}}", noticia.get("link", ""))

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        prs.save(temp_file.name)

        return send_file(temp_file.name, as_attachment=True, download_name="noticia.pptx")
    except Exception as e:
        print("🔥 Erro:", str(e))
        return {"erro": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
