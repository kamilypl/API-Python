from flask import Flask, request, send_file
from pptx import Presentation
from pptx.util import Pt
import tempfile
import os

# Cria a aplicação Flask
app = Flask(__name__)

# Define o caminho do template PowerPoint (template.pptx) baseado no diretório atual do script
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

# Rota que aceita POST para gerar o PPTX com as notícias
@app.route("/gerar_pptx", methods=["POST"])
def gerar_pptx():
    try:
        # Lê o corpo da requisição e extrai a lista de notícias
        dados = request.json
        noticias = dados.get("noticias", [])

        # Carrega o template PPTX
        prs = Presentation(TEMPLATE_PATH)
        slide = prs.slides[0]

        preenchidos = 0  # Contador de caixas preenchidas

        # Itera pelas shapes do slide
        for shape in slide.shapes:
            if not shape.has_text_frame or preenchidos >= len(noticias):
                continue

            noticia = noticias[preenchidos]

            for paragraph in shape.text_frame.paragraphs:
                # Junta todo o texto do parágrafo, incluindo runs separados
                texto_total = "".join(run.text for run in paragraph.runs)

                # Substitui os placeholders pelos dados reais
                texto_formatado = (
                    texto_total
                    .replace("{{titulo}}", noticia.get("titulo", ""))
                    .replace("{{resumo}}", "\n" + noticia.get("resumo", ""))
                    .replace("{{data}}", "\nData: " + noticia.get("data", ""))
                    .replace("{{link}}", "\n" + noticia.get("link", ""))
                )

                # Remove todos os runs antigos do parágrafo
                while paragraph.runs:
                    paragraph._element.remove(paragraph.runs[0]._element)

                # Cria um novo run e define o texto formatado
                run = paragraph.add_run()
                run.text = texto_formatado
                run.font.size = Pt(12)  # Ajuste de tamanho conforme o template

            preenchidos += 1  # Próxima notícia

        print(f"✅ Blocos preenchidos: {preenchidos}")

        # Salva o arquivo final como temporário
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        prs.save(temp_file.name)

        # Retorna o arquivo como resposta
        return send_file(
            temp_file.name,
            as_attachment=True,
            download_name="noticias_atualizadas.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        print("🔥 Erro interno:", str(e))
        return {"erro": str(e)}, 500

# Roda o servidor Flask na porta 10000
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
