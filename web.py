import os
import io
from flask import Flask, request, redirect, url_for, render_template_string, send_file, abort
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import Workbook

# Configuration via environment variables:
# - GOOGLE_SERVICE_ACCOUNT_JSON: path to service account JSON file (or set GOOGLE_APPLICATION_CREDENTIALS)
# - SHEET_ID: Google Sheets ID (from the URL)
# - ADMIN_PASSWORD: password for admin panel

SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "service_account.json")
SHEET_ID = os.getenv("SHEET_ID")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "changeme")

if not SHEET_ID:
    raise RuntimeError("SHEET_ID environment variable not set. See README_GOOGLE.md for setup.")

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPE)
client = gspread.authorize(creds)
sh = client.open_by_key(SHEET_ID)
ws = sh.sheet1

app = Flask(__name__)

FORM_HTML = '''
<!doctype html>
<html>
<head>
  <title>Rifa Fegsat - Registrar Venda</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      margin: 0;
      padding: 20px;
      min-height: 100vh;c
    }
    .container {
      max-width: 500px;
      margin: 0 auto;
      background: white;
      border-radius: 12px;
      padding: 30px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    }
    h1 {
      color: #333;
      text-align: center;
      margin-bottom: 30px;
    }
    .form-group {
      margin-bottom: 15px;
    }
    label {
      display: block;
      margin-bottom: 5px;
      color: #555;
      font-weight: bold;
    }
    input[type="text"],
    input[type="password"] {
      width: 100%;
      padding: 10px;
      border: 1px solid #ddd;
      border-radius: 8px;
      font-size: 14px;
      box-sizing: border-box;
    }
    input[type="submit"],
    .button {
      width: 100%;
      padding: 12px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: bold;
      cursor: pointer;
      margin-top: 10px;
    }
    input[type="submit"]:hover,
    .button:hover {
      transform: scale(1.02);
      box-shadow: 0 6px 12px rgba(0, 0, 0, 0.2);
    }
    .links {
      text-align: center;
      margin-top: 20px;
    }
    .links a {
      display: inline-block;
      margin: 10px 5px;
      padding: 10px 20px;
      background: #f0f0f0;
      color: #667eea;
      text-decoration: none;
      border-radius: 8px;
      font-weight: bold;
      transition: all 0.3s;
    }
    .links a:hover {
      background: #667eea;
      color: white;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>🎫 Registrar Venda</h1>
    <form method=post action="/submit">
      <div class="form-group">
        <label>Nome:</label>
        <input type=text name=nome required>
      </div>
      <div class="form-group">
        <label>Telefone:</label>
        <input type=text name=telefone required>
      </div>
      <div class="form-group">
        <label>Número:</label>
        <input type=text name=numero required>
      </div>
      <input type=submit value="Registrar">
    </form>
    <div class="links">
      <a href="/sold">Ver Números Vendidos</a>
      <a href="/admin">Painel Admin</a>
    </div>
  </div>
</body>
</html>
'''

SOLD_HTML = '''
<!doctype html>
<html>
<head>
  <title>Números Vendidos - Rifa Fegsat</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      margin: 0;
      padding: 20px;
      min-height: 100vh;
    }
    .container {
      max-width: 900px;
      margin: 0 auto;
      background: white;
      border-radius: 12px;
      padding: 30px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    }
    h1 {
      color: #333;
      text-align: center;
      margin-bottom: 10px;
    }
    .stats {
      text-align: center;
      font-size: 18px;
      color: #666;
      margin-bottom: 30px;
      padding: 10px;
      background: #f5f5f5;
      border-radius: 8px;
    }
    .numbers-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(80px, 1fr));
      gap: 10px;
      margin-bottom: 30px;
    }
    .number-badge {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 15px;
      border-radius: 8px;
      text-align: center;
      font-weight: bold;
      font-size: 18px;
      box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .no-data {
      text-align: center;
      color: #999;
      font-size: 16px;
      padding: 40px;
    }
    .button {
      display: inline-block;
      padding: 12px 24px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      text-decoration: none;
      border-radius: 8px;
      text-align: center;
      cursor: pointer;
      font-size: 16px;
    }
    .button:hover {
      transform: scale(1.05);
      box-shadow: 0 6px 12px rgba(0, 0, 0, 0.2);
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>🎫 Números Vendidos</h1>
    <div class="stats">
      <strong>Total de números vendidos: {{total}}</strong>
    </div>
    {{content}}
    <div style="text-align: center; margin-top: 20px;">
      <a href="/" class="button">← Voltar</a>
    </div>
  </div>
</body>
</html>
'''

ADMIN_LOGIN_HTML = '''
<!doctype html>
<html>
<head>
  <title>Admin Login - Rifa Fegsat</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      margin: 0;
      padding: 20px;
      min-height: 100vh;
      display: flex;
      justify-content: center;
      align-items: center;
    }
    .container {
      background: white;
      border-radius: 12px;
      padding: 40px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
      width: 100%;
      max-width: 400px;
    }
    h1 {
      color: #333;
      text-align: center;
      margin-bottom: 30px;
    }
    .form-group {
      margin-bottom: 20px;
    }
    label {
      display: block;
      margin-bottom: 8px;
      color: #555;
      font-weight: bold;
    }
    input[type="password"] {
      width: 100%;
      padding: 12px;
      border: 1px solid #ddd;
      border-radius: 8px;
      font-size: 16px;
      box-sizing: border-box;
    }
    input[type="submit"] {
      width: 100%;
      padding: 12px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: bold;
      cursor: pointer;
    }
    input[type="submit"]:hover {
      transform: scale(1.02);
      box-shadow: 0 6px 12px rgba(0, 0, 0, 0.2);
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>🔐 Administrador</h1>
    <form method=post action="/admin">
      <div class="form-group">
        <label>Senha:</label>
        <input type=password name=password required autofocus>
      </div>
      <input type=submit value="Entrar">
    </form>
  </div>
</body>
</html>
'''

ADMIN_PANEL_HTML = '''
<!doctype html>
<html>
<head>
  <title>Painel Admin - Rifa Fegsat</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      margin: 0;
      padding: 20px;
      min-height: 100vh;
    }
    .container {
      max-width: 600px;
      margin: 0 auto;
      background: white;
      border-radius: 12px;
      padding: 30px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
    }
    h1 {
      color: #333;
      text-align: center;
      margin-bottom: 30px;
    }
    .button-group {
      display: flex;
      gap: 15px;
      margin-bottom: 30px;
    }
    .button-group form {
      flex: 1;
    }
    button, input[type="submit"] {
      width: 100%;
      padding: 15px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: bold;
      cursor: pointer;
      transition: all 0.3s;
    }
    button:hover, input[type="submit"]:hover {
      transform: scale(1.05);
      box-shadow: 0 6px 12px rgba(0, 0, 0, 0.2);
    }
    .danger {
      background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    }
    .back-link {
      text-align: center;
      margin-top: 20px;
    }
    .back-link a {
      color: #667eea;
      text-decoration: none;
      font-weight: bold;
    }
    .back-link a:hover {
      text-decoration: underline;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>⚙️ Painel Admin</h1>
    <div class="button-group">
      <form method=post action="/admin/clear" onsubmit="return confirm('Tem certeza? Esta ação é irreversível!');">
        <input type=hidden name=password value="{{password}}">
        <input type=submit value="🗑️ Apagar Todas as Vendas" class="danger">
      </form>
      <form method=get action="/admin/export">
        <input type=submit value="📥 Exportar XLSX">
      </form>
    </div>
    <div class="back-link">
      <a href="/">← Voltar</a>
    </div>
  </div>
</body>
</html>
'''

@app.route('/', methods=['GET'])
def index():
    return render_template_string(FORM_HTML)

@app.route('/submit', methods=['POST'])
def submit():
    nome = request.form.get('nome','').strip()
    telefone = request.form.get('telefone','').strip()
    numero = request.form.get('numero','').strip()

    if not nome or not telefone or not numero:
        return "Preencha todos os campos", 400

    # Check duplicate numero
    rows = ws.get_all_values()[1:]  # skip header
    for r in rows:
        if len(r) >= 3 and r[2] == numero:
            return f"O número {numero} já foi vendido", 400

    ws.append_row([nome, telefone, numero])
    return redirect(url_for('index'))

@app.route('/admin', methods=['GET','POST'])
def admin():
    if request.method == 'GET':
        return render_template_string(ADMIN_LOGIN_HTML)

    password = request.form.get('password','')
    if password != ADMIN_PASSWORD:
        return "Senha incorreta", 403

    # render admin panel with hidden password so admin_clear can verify
    return render_template_string(ADMIN_PANEL_HTML, password=password)

@app.route('/admin/clear', methods=['POST'])
def admin_clear():
    password = request.form.get('password') or request.args.get('password')

    if not password or password != ADMIN_PASSWORD:
        return "Senha inválida para operação administrativa.", 403

    # Attempt to clear the sheet, report errors if any
    try:
        header = ws.row_values(1)
        # Delete only data rows, keep header
        rows = ws.get_all_values()
        if len(rows) > 1:
            # delete_rows is exclusive at the end, so add 1 to include the last row
            ws.delete_rows(2, len(rows) + 1)
        # ensure header exists
        if not header:
            ws.append_row(["Nome","Telefone","Numero"])
    except Exception as e:
        # log error and return readable message
        return f"Erro ao apagar dados: {e}", 500

    return redirect(url_for('admin'))

@app.route('/admin/export', methods=['GET'])
def admin_export():
    # Read all values and write XLSX in-memory
    values = ws.get_all_values()
    wb = Workbook()
    w = wb.active
    for row in values:
        w.append(row)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    return send_file(bio,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,
                     download_name='dados.xlsx')

@app.route('/sold', methods=['GET'])
def sold():
    rows = ws.get_all_values()[1:]  # skip header
    sold_numbers = [row[2] for row in rows if len(row) >= 3 and row[2].strip()]
    
    if sold_numbers:
        content = '<div class="numbers-grid">'
        for num in sorted(sold_numbers, key=lambda x: int(x) if x.isdigit() else 0):
            content += f'{num}'
        content += '</div>'
    else:
        content = 'Nenhum número vendido ainda.'
    
    return render_template_string(SOLD_HTML, total=len(sold_numbers), content=content)

if __name__ == '__main__':
    # For development only. In production use gunicorn/uwsgi
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
