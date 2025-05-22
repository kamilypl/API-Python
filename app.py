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
                # Junta todo o conteúdo dos runs
                texto_total = "".join(run.text for run in paragraph.runs)

                # Faz as substituições
                texto_formatado = (
                    texto_total
                    .replace("{{titulo}}", noticia.get("titulo", ""))
                    .replace("{{resumo}}", "\n" + noticia.get("resumo", ""))
                    .replace("{{data}}", "\nData: " + noticia.get("data", ""))
                    .replace("{{link}}", "\n" + noticia.get("link", ""))
                )

                # Limpa os runs e insere o texto formatado no primeiro run
                for run in paragraph.runs:
                    run.text = ""
                if paragraph.runs:
                    paragraph.runs[0].text = texto_formatado

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
