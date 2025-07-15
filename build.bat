@echo off
echo Suspect Word Edit - Script de Compilação
echo ==============================
echo.

REM Verifica se o Python ta instalado
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo Erro: Python nao encontrado bixo burro. instala o Python e tenta novamente.
    pause
    exit /b 1
)

REM Verifica se o arquivo icon.ico existe
if not exist icon.ico (
    echo Erro: Arquivo icon.ico nao encontrado. 
    echo Este arquivo e necessario para a compilacao.
    echo Por favor, certifique-se de que o arquivo icon.ico esta na pasta do projeto.
    pause
    exit /b 1
)

REM Instala ou atualiza as dependencias
echo Instalando as dependencias...
pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo Erro ao instalar as dependencias.
    pause
    exit /b 1
)

REM Limpar diretórios de compilação anteriores
echo Limpando compilacoes anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Compila
echo Compilando o programa...
echo Usando o icone: icon.ico
pyinstaller suspect.spec
if %ERRORLEVEL% NEQ 0 (
    echo Erro durante a compilacao.
    pause
    exit /b 1
)

echo.
echo Compilacao concluida com sucesso krl amem
echo O executavel ta disponivel na pasta 'dist'
echo.
pause 