name: Generar Folletos Mensuales

on:
  # Disparado manualmente desde admin.html via GitHub API
  repository_dispatch:
    types: [generar-folletos]

  # También permite disparo manual desde GitHub UI
  workflow_dispatch:
    inputs:
      target_month:
        description: 'Mes objetivo YYYY-MM (vacío = mes anterior automático)'
        required: false
        default: ''

jobs:
  generar:
    runs-on: ubuntu-latest
    timeout-minutes: 30

    steps:
      - name: Checkout repo
        uses: actions/checkout@v4
        with:
          token: ${{ secrets.GITHUB_TOKEN }}
          fetch-depth: 1

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Instalar dependencias
        run: |
          pip install -r scripts/requirements.txt

      - name: Instalar Playwright Chromium
        run: |
          playwright install chromium --with-deps

      - name: Instalar ODBC Driver (para SQL Server)
        run: |
          curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add -
          curl https://packages.microsoft.com/config/ubuntu/22.04/prod.list \
            > /etc/apt/sources.list.d/mssql-release.list
          apt-get update -qq
          ACCEPT_EULA=Y apt-get install -y msodbcsql17 unixodbc-dev
        continue-on-error: true

      - name: Generar folletos
        env:
          SQL_CONN_STR:  ${{ secrets.SQL_CONN_STR }}
          TARGET_MONTH:  ${{ github.event.inputs.target_month || github.event.client_payload.target_month || '' }}
        run: |
          cd scripts
          python main.py

      - name: Commit folletos generados
        run: |
          git config user.name  "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add folletos/ scripts/etl/cmf_scraper.py
          git diff --cached --quiet || git commit -m "🗂 Folletos $(date +'%Y-%m') generados automáticamente"
          git push
