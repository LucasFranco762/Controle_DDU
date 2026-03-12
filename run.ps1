# Script para executar o Controle_DDU ativando o venv automaticamente
param(
    [switch]$NoWait = $false
)

Write-Host "Ativando ambiente virtual..." -ForegroundColor Green
& .\.venv\Scripts\Activate.ps1

Write-Host "Iniciando Controle_DDU..." -ForegroundColor Green

if ($NoWait) {
    Start-Process python -ArgumentList "Controle_documentos.py" -NoNewWindow
} else {
    python Controle_documentos.py
}
