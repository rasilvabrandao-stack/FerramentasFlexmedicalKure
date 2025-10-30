import requests # type: ignore
import json

WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzka9zfxb9UcVz2kVafIWmiYT12YHx0JPb3zPU8jU1PN4BuNzBXeVUe1bMxxqG21b6O0A/exec"  # URL do deploy do Apps Script

# Payload de teste para movimentação de ferramenta
movimentacao = {
    "tipo": "movimentacao",
    "solicitante": "BRUNO GOMES DA SILVA",
    "tipoMov": "retirada",
    "ferramenta": "Furadeira",
    "patrimonio": "PAT001",

    "dataSaida": "2024-01-15",
    "horaSaida": "08:30",
    "dataRetorno": "",
    "temRetorno": "sim",
    "observacao": ""
}

response = requests.post(
    WEBAPP_URL,
    headers={"Content-Type": "application/json"},
    data=json.dumps(movimentacao)
)

print(response.text)
