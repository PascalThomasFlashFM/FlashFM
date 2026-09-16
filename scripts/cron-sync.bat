@echo off
REM Appele par le Planificateur de taches Windows. Ne declenche jamais d'envoi
REM de mail : la validation des confirmations reste manuelle (serve / send).
cd /d "%~dp0.."
node dist\cli.js sync >> data\cron.log 2>&1
