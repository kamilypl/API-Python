# API Geradora de PowerPoint (.pptx) para Automatização

## Como rodar localmente

1. Instale as dependências:
   pip install -r requirements.txt

2. Execute:
   python app.py

Acesse http://localhost:10000/gerar_pptx

## Como subir no Render.com

1. Faça login no Render.com e crie um novo "Web Service".
2. Conecte seu repositório GitHub com os arquivos acima.
3. Configure:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
   - Environment: Python 3.10+
   - Porta: Render detecta automaticamente (use a variável $PORT)
4. Faça upload do arquivo template.pptx no mesmo diretório.

Pronto! A API ficará disponível e você poderá chamá-la pelo N8N via HTTP POST.

## Exemplo de chamada via HTTP POST

Endpoint: `/gerar_pptx`

Payload:
```json
{
  "titulo": "Mercado de Ar-condicionado cresce 12%",
  "resumo": "Novo estudo aponta expansão do setor em 2025...",
  "link": "https://link-da-noticia",
  "data": "16/05/2025",
  "setor": "Ar-condicionado"
}
#   A P I - P y t h o n  
 #   A P I - P y t h o n  
 