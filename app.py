from flask import Flask, request, send_file
from pptx import Presentation
import tempfile
import os

# Cria a aplicação Flask
app = Flask(__name__)

# Define o caminho do template PowerPoint (template.pptx) baseado no diretório atual do script
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.pptx')

# Define uma rota chamada "/gerar_pptx" que aceita requisições do tipo POST
@app.route("/gerar_pptx", methods=["POST"])
def gerar_pptx():
    try:
        # Lê o corpo JSON da requisição recebida
        dados = request.json

        # Extrai a lista de notícias do dicionário (ou uma lista vazia se não houver "noticias")
        noticias = dados.get("noticias", [])

        # Carrega a apresentação base (template)
        prs = Presentation(TEMPLATE_PATH)

        # Seleciona o primeiro slide do template
        slide = prs.slides[0]

        # Variável para contar quantas caixas de texto foram preenchidas
        preenchidos = 0

        # Itera por todas as formas (shapes) do slide
        for shape in slide.shapes:
            # Pula se o shape não tem quadro de texto ou se já preencheu todas as notícias
            if not shape.has_text_frame or preenchidos >= len(noticias):
                continue

            # Pega a notícia correspondente ao índice atual
            noticia = noticias[preenchidos]

            # Itera pelos parágrafos do quadro de texto
            for paragraph in shape.text_frame.paragraphs:
                # Junta todo o conteúdo dos "runs" (partes de texto com formatação individual)
                texto_total = "".join(run.text for run in paragraph.runs)

                # Substitui os marcadores com os dados da notícia
                texto_formatado = (
                    texto_total
                    .replace("{{titulo}}", noticia.get("titulo", ""))
                    .replace("{{resumo}}", "\n" + noticia.get("resumo", ""))
                    .replace("{{data}}", "\nData: " + noticia.get("data", ""))
                    .replace("{{link}}", "\n" + noticia.get("link", ""))
                )

                # Limpa o texto de todos os runs
                for run in paragraph.runs:
                    run.text = ""

                # Atribui o texto formatado ao primeiro run, se houver
                if paragraph.runs:
                    paragraph.runs[0].text = texto_formatado

            # Incrementa o contador de blocos preenchidos
            preenchidos += 1

        # Exibe no console o número de blocos preenchidos com sucesso
        print(f"✅ Blocos preenchidos: {preenchidos}")

        # Cria um arquivo temporário com extensão .pptx
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")

        # Salva a apresentação modificada nesse arquivo temporário
        prs.save(temp_file.name)

        # Envia o arquivo gerado como resposta para download
        return send_file(
            temp_file.name,
            as_attachment=True,  # Faz o navegador baixar o arquivo
            download_name="noticias_atualizadas.pptx",  # Nome sugerido para o arquivo baixado
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"  # Tipo MIME de arquivos PowerPoint
        )

    except Exception as e:
        print("🔥 Erro interno:", str(e))
        return {"erro": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
