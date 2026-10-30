@echo off
chcp 65001 >nul
title Configurare Windows Firewall - Acces Retea Locala (Port 8080-8090)

:: Verifica daca ruleaza cu drepturi de Administrator; daca nu, cere automat permisiunea UAC
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Se solicita drepturi de Administrator pentru a deschide portul 8080 in Windows Firewall...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo ========================================================================
echo   Activare Acces in Reteaua Locala (LAN / Wi-Fi) pentru Platforma Web
echo ========================================================================

netsh advfirewall firewall delete rule name="Platforma Analiza Financiara (LAN)" >nul 2>&1
netsh advfirewall firewall add rule name="Platforma Analiza Financiara (LAN)" dir=in action=allow protocol=TCP localport=8080-8090 profile=private,domain,public

echo.
echo [OK] Porturile TCP 8080-8090 au fost deschise in Windows Firewall!
echo      Acum colegii din reteaua locala pot accesa site-ul folosind adresa IP afisata.
echo.
pause

