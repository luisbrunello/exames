import io
import zipfile
from flask import Flask, Response, request, render_template_string

from exam_processor import process_files, report_to_json_bytes

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB

HTML = """
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Preenchimento de Exames</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background: #f6f7fb; color: #1c2430; margin: 0; }
    .wrap { max-width: 760px; margin: 48px auto; padding: 24px; }
    .card { background: white; border-radius: 18px; padding: 28px; box-shadow: 0 10px 30px rgba(0,0,0,.08); }
    h1 { margin-top: 0; font-size: 28px; }
    p { line-height: 1.55; color: #4c5a6a; }
    label { display: block; margin: 18px 0 8px; font-weight: 600; }
    input[type=file] { width: 100%; padding: 14px; border: 1px solid #d7dce5; border-radius: 12px; background: #fbfcfe; }
    button { margin-top: 22px; background: #162d47; color: white; border: 0; border-radius: 12px; padding: 14px 18px; font-size: 15px; cursor: pointer; }
    button:hover { opacity: .95; }
    .hint { font-size: 14px; color: #667587; margin-top: 10px; }
    .error { background: #fff2f2; color: #a12b2b; border: 1px solid #f2c9c9; border-radius: 12px; padding: 12px 14px; margin-bottom: 16px; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>Preencher painel de exames</h1>
      <p>Envie o PDF do laboratório e a planilha Excel. O sistema adiciona uma nova coluna no próximo espaço disponível, escreve a data na linha <strong>DATA</strong> e mantém as colunas antigas intactas.</p>
      {% if error %}<div class="error">{{ error }}</div>{% endif %}
      <form method="post" enctype="multipart/form-data">
        <label for="pdf_file">PDF do exame</label>
        <input id="pdf_file" type="file" name="pdf_file" accept=".pdf" required>

        <label for="xlsx_file">Planilha Excel</label>
        <input id="xlsx_file" type="file" name="xlsx_file" accept=".xlsx" required>

        <div class="hint">O download será um arquivo ZIP contendo a planilha atualizada e um relatório JSON.</div>
        <button type="submit">Processar arquivos</button>
      </form>
    </div>
  </div>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(HTML, error=None)


@app.post("/")
def process_upload():
    pdf_file = request.files.get("pdf_file")
    xlsx_file = request.files.get("xlsx_file")

    if not pdf_file or not pdf_file.filename:
        return render_template_string(HTML, error="Envie o PDF do exame.")
    if not xlsx_file or not xlsx_file.filename:
        return render_template_string(HTML, error="Envie a planilha Excel.")
    if not pdf_file.filename.lower().endswith(".pdf"):
        return render_template_string(HTML, error="O arquivo do exame precisa ser .pdf.")
    if not xlsx_file.filename.lower().endswith(".xlsx"):
        return render_template_string(HTML, error="A planilha precisa ser .xlsx.")

    try:
        workbook_stream, report = process_files(pdf_file.stream, xlsx_file.stream, pdf_file.filename, xlsx_file.filename)
    except Exception as exc:
        return render_template_string(HTML, error=f"Erro ao processar os arquivos: {exc}")

    zip_buffer = io.BytesIO()
    output_xlsx_name = xlsx_file.filename.rsplit(".", 1)[0] + "_atualizado.xlsx"
    report_name = "relatorio_extracao.json"

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(output_xlsx_name, workbook_stream.getvalue())
        zf.writestr(report_name, report_to_json_bytes(report))

    zip_buffer.seek(0)

    return Response(
        zip_buffer.getvalue(),
        mimetype="application/zip",
        headers={
            "Content-Disposition": "attachment; filename=resultado_exames.zip"
        },
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
