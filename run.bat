@echo off
REM Удобный запуск системы сверки БУ-ПУ (Windows).
REM Без аргументов открывает интерактивное меню; иначе пробрасывает флаги в CLI.
cd /d "%~dp0"

if "%~1"=="" (
  python -m npf_recon --menu
) else (
  python -m npf_recon %*
)
