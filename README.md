# Movimientos de equipos — Add-in MyGeotab

Registro de **instalaciones, desinstalaciones y reemplazos** que los clientes hacen por su cuenta. Cada movimiento se aplica en MyGeotab y queda registrado en `AddInData`.

## Qué hace en MyGeotab
| Movimiento | Acción sobre `Device` |
|---|---|
| Instalación, patente nueva | `Add` con el grupo elegido |
| Instalación, patente archivada | `Set`: reactiva (`activeTo` = 2050) y asigna el GO |
| Instalación, patente activa sin GO | `Set`: asigna el GO |
| Instalación, patente activa con GO | Se rechaza: usar Reemplazo |
| Reemplazo de GO | `Set`: cambia `serialNumber` y `deviceType` en el mismo vehículo |
| Reemplazo de cámara o arnés | Sin cambios en `Device` |
| Desinstalación de GO | `Set`: archiva (`activeTo` = ahora). Nunca elimina. Geotab deja de facturarlo |
| Desinstalación solo de cámara | Sin cambios en `Device` |

El odómetro se corrige con `Add StatusData` (`DiagnosticOdometerAdjustmentId`) solo si difiere de la lectura actual en 1 km o más.
Si Geotab rechaza el cambio no se registra el movimiento. Anular un movimiento no revierte MyGeotab.

## Vehículos
Lista todos los vehículos con su primera instalación: la del add-in si existe; si no, `activeFrom`.

## Instalación
1. Publicar el repo en GitHub Pages (*Settings → Pages → Branch: main, carpeta /root*).
2. Pegar `dist/config-hosted.json` en *Administración → Sistema → Configuración del sistema → Add-Ins*.
3. Queda en *Actividad → Movimientos de equipos*.

Para otra URL de hosting: `python build.py URL_BASE`.

## Estado de comunicación
Se abre al registrar una instalación o un reemplazo de GO, y desde *Historial → Comunicación*. **Refrescar estado** vuelve a consultar Geotab:
alta del GO, comunicación (`DeviceStatusInfo`), posiciones (`LogRecord`), RPM (`StatusData`) y VIN del motor contra el chasis declarado.
Solo cuentan registros reales desde que el add-in asignó el GO (o desde el día del movimiento en registros anteriores). La consulta no se guarda ni cubre la cámara.

## Permisos
- Registrar: además de datos de add-ins, el usuario necesita permiso para crear y editar vehículos y para escribir datos de estado.
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
