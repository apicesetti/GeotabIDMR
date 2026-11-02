# Movimientos de equipos — Add-in MyGeotab

Registro de **instalaciones, desinstalaciones y reemplazos** que los clientes hacen por su cuenta.
Solo escribe en `AddInData`: **no modifica vehículos ni equipos en Geotab**.

## Instalación
1. Publicar el repo en GitHub Pages (*Settings → Pages → Branch: main, carpeta /root*).
2. Pegar `dist/config-hosted.json` en *Administración → Sistema → Configuración del sistema → Add-Ins*.
3. Queda en *Actividad → Movimientos de equipos*.

Para otra URL de hosting: `python build.py URL_BASE`.

## Permisos
- Registrar y consultar: usuarios con acceso de escritura a datos de add-ins.
- Anular: solo perfil **Administrador** de MyGeotab (`GroupEverythingSecurityId`). El control está en la interfaz del add-in; la API de MyGeotab no lo impone sobre `AddInData`.

## Modelo de datos (`AddInData`, `addInId = aOS_yCVsbRPiyLHnQhiSteg`)
Un registro por movimiento, `details.type = "movimiento"`:

| Campo | Contenido |
|---|---|
| `movementType` | `instalacion`, `desinstalacion`, `reemplazo` |
| `date`, `odometerKm`, `notes` | Fecha declarada, km y observaciones |
| `deviceId`, `deviceName`, `licensePlate`, `chassis`, `vehicleBrand`, `vehicleModel` | Vehículo |
| `goSerial`, `harness`, `camera {model, serial}` | Equipo |
| `removed` | Desinstalación: `go`, `go_camara`, `camara` |
| `replaced {go, camera, harness}` | Reemplazo: cada uno con `from` y `to` |
| `createdAt`, `createdBy` | Quién y cuándo cargó |
| `voided`, `voidReason`, `voidedBy`, `voidedAt` | Anulación (el registro no se borra) |

## Consulta por API (Python)
```python
recs = api.get("AddInData", search={"addInId": "aOS_yCVsbRPiyLHnQhiSteg"})
movs = [r["details"] for r in recs
        if r["details"].get("type") == "movimiento" and not r["details"].get("voided")]
```
