#!/usr/bin/env bash
set -euo pipefail

# Build script for PyInstaller. Usage: ./build.sh
APP="main.py"
APP_NAME="Rifa Fegsat"
DATAFILE="dados.xlsx"
ICON_ICO="/Users/joao_eberle/Downloads/Logo_app_rifa.ico"
ICON_ICNS="/Users/joao_eberle/Downloads/icon.icns"
ICON=""

# choose icon based on platform and available files
if [[ "${OSTYPE}" == "darwin"* ]]; then
  if [[ -f "$ICON_ICNS" ]]; then
    ICON="$ICON_ICNS"
  elif [[ -f "$ICON_ICO" ]]; then
    ICON="$ICON_ICO"
  fi
else
  if [[ -f "$ICON_ICO" ]]; then
    ICON="$ICON_ICO"
  elif [[ -f "$ICON_ICNS" ]]; then
    ICON="$ICON_ICNS"
  fi
fi

# Ensure pyinstaller is installed: pip install -r requirements.txt

if [[ "${OSTYPE}" == "msys" || "${OSTYPE}" == "win32" ]]; then
  # Windows (Git Bash/MSYS) use semicolon separator
  pyinstaller --onedir --noconfirm --noconsole --name "$APP_NAME" --add-data "${DATAFILE};." ${ICON:+--icon "$ICON"} "$APP"
else
  # macOS / Linux use colon separator
  pyinstaller --onedir --noconfirm --noconsole --name "$APP_NAME" --add-data "${DATAFILE}:." ${ICON:+--icon "$ICON"} "$APP"
fi

echo "Build finished. Check the dist/ folder for the application."