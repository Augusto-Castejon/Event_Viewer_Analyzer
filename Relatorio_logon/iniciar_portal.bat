:: Arquivo: iniciar_portal.bat
@echo off
powershell -ExecutionPolicy Bypass -File "C:\Temp\Relatorio_logon\Relatorio_RDP_Filtravel.ps1"
powershell -ExecutionPolicy Bypass -File "C:\Temp\Relatorio_logon\Relatorio_Desligamento.ps1"
powershell -ExecutionPolicy Bypass -File "C:\Temp\Relatorio_logon\Relatorio_Modificacoes_Arquivos.ps1"
powershell -ExecutionPolicy Bypass -File "C:\Temp\Relatorio_logon\Relatorio_Instalacao_Softwares.ps1"
start "" "C:\Temp\Relatorio_logon\index.html"