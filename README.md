# reporte-parking

`index.html` se actualiza solo, todos los dias (~08:00 Chile), via GitHub Actions
(`.github/workflows/actualizar-reporte.yml`). El script (`scripts/actualizar_reporte.py`)
descarga los pagos nuevos de ParkingApp, los reconcilia contra `data/pagos.json`
(historial completo de transacciones) y regenera el HTML desde `scripts/template.html`.

Requiere dos secrets del repo (Settings → Secrets and variables → Actions):
`PARKINGAPP_EMAIL` y `PARKINGAPP_PASSWORD`.

El Excel `Parking Switch 345.xlsm` (Módulo5) sigue funcionando igual que antes,
de forma manual/independiente - no comparte datos con este script.