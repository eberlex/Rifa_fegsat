Configurar Google Sheets (service account)

1. Criar service account no Google Cloud Console
   - IAM & Admin > Service Accounts > Create
   - Criar e gerar uma chave JSON; baixe o arquivo
2. Compartilhar a planilha com o e-mail do service account (xxxxx@...gserviceaccount.com) com permissão Editor
3. Anotar o ID da planilha da URL (entre /d/ and /edit)
4. Copie o JSON para o servidor e defina a variável de ambiente:
   export GOOGLE_SERVICE_ACCOUNT_JSON=/path/to/service_account.json
   export SHEET_ID=your_sheet_id_here
   export ADMIN_PASSWORD=uma_senha_segura
5. Rodar localmente (dev):
   pip install -r requirements.txt
   python web.py

Notas:
- Em produção, usar gunicorn ou outro WSGI server e não executar o Flask dev server direto.
- Se preferir usar Google Application Default Credentials, configure GOOGLE_APPLICATION_CREDENTIALS em vez de GOOGLE_SERVICE_ACCOUNT_JSON.
