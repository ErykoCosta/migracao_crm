@echo off
rem Abre o Chrome com porta de depuracao 9222 e perfil dedicado (a sessao do login fica salva).
set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" (
  echo Chrome nao encontrado. Edite este arquivo e ajuste o caminho em CHROME.
  pause
  exit /b 1
)
start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="C:\chrome_automacao" https://app.advbox.com.br/

