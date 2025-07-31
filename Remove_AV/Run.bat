@echo off 

echo Starting ....

powershell.exe -executionpolicy bypass -file "%~dp0\McAfee-Remover.ps1"

echo Done 

