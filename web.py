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
<title>Rifa Fegsat - Registrar Venda</title>
<h1>Registrar Venda</h1>
<form method=post action="/submit">
  Nome: <input type=text name=nome required><br>
  Telefone: <input type=text name=telefone required><br>
  Número: <input type=text name=numero required><br>
  <input type=submit value="Registrar">
</form>
<p><a href="/sold">Ver Números Vendidos</a></p>
<p><a href="/admin">Painel Admin</a></p>
'''

SOLD_HTML = '''
<!doctype html>
<title>Números Vendidos - Rifa Fegsat</title>
<h1>Números Vendidos</h1>
<p>Total de números vendidos: {{total}}</p>
{{content}}
<p><a href="/">Voltar</a></p>
'''

ADMIN_LOGIN_HTML = '''
<!doctype html>
<title>Admin Login</title>
<h1>Administrador</h1>
<form method=post action="/admin">
  Senha: <input type=password name=password required>
  <input type=submit value="Entrar">
</form>
'''

ADMIN_PANEL_HTML = '''
<!doctype html>
<title>Painel Admin</title>
<h1>Painel Admin</h1>
<!-- password is injected as hidden field when rendering -->
<form method=post action="/admin/clear" onsubmit="return confirm('Apagar todas as vendas?');">
  <input type=hidden name=password value="{{password}}">
  <input type=submit value="Apagar Todas as Vendas">
</form>
<form method=get action="/admin/export">
  <input type=submit value="Exportar XLSX">
</form>
<p><a href="/">Voltar</a></p>
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
        content = '<ul>'
        for num in sorted(sold_numbers, key=lambda x: int(x) if x.isdigit() else 0):
            content += f'<li>{num}</li>'
        content += '</ul>'
    else:
        content = '<p>Nenhum número vendido ainda.</p>'
    
    return render_template_string(SOLD_HTML, total=len(sold_numbers), content=content)

if __name__ == '__main__':
    # For development only. In production use gunicorn/uwsgi
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
