@echo off
title Tidrapportering

:: ============================================================
:: Steg 1 – Generera Excel-underlagen
:: ============================================================
echo.
echo [1/3] Genererar Excel-underlagen...
echo.
call run.bat
echo.
echo Klart! Excelfiler sparade i data\.
echo.
pause

:: ============================================================
:: Steg 2 – Starta Chrome med Antura
:: ============================================================
echo.
echo [2/3] Startar Chrome...
echo.
start "" cmd /c start_chrome.bat
echo.
echo Webbläsaren är öppen och Antura bör visas.
echo Logga in med ditt UCR-konto om det behövs (behövs normalt bara första gången).
echo.
pause

:: ============================================================
:: Steg 3 – Fyll i Antura automatiskt
:: ============================================================
echo.
echo [3/3] Startar Antura-fyllaren...
echo.
call run_antura.bat
