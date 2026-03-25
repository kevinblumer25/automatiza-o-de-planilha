@echo off
cd /d "%~dp0"

echo Executando o script, aguarde...
python main.py
set exit_code=%ERRORLEVEL%

if %exit_code% EQU 0 (
	echo Concluído com sucesso.
) else (
	echo Erro na execução. Código: %exit_code%
)

exit /b %exit_code%