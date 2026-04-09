# Plataforma de Pólizas de Ingreso — RZ2 Sistemas
## GBC Business Consulting

---

## Instalación (una sola vez)

### Requisitos previos
- Python 3.9 o superior instalado en la computadora

### Windows
1. Doble clic en `INICIAR_WINDOWS.bat`
2. La primera vez instala las dependencias automáticamente
3. Se abre la plataforma en tu navegador en http://localhost:5050

### Mac
1. Abre Terminal en la carpeta
2. Ejecuta: `bash INICIAR_MAC.command`
3. Abre tu navegador en http://localhost:5050

---

## Uso mensual (contador junior)

### Archivos que necesitas cada mes:
| Archivo | Formato | Descripción |
|---------|---------|-------------|
| Extracto bancario | `.csv` | Exportado del portal Scotiabank |
| Facturas pendientes | `.xlsx` | Hoja `2026` con facturas vigentes |
| Catálogo de cuentas | `.xlsx` | **Opcional** — hoja `cuentas` del archivo de pólizas anterior |

### Proceso:
1. Ejecuta el script de inicio
2. Sube los archivos en la pantalla (arrastra o clic)
3. Haz clic en "Generar pólizas CONTPAq"
4. Descarga el Excel resultante

### El Excel de salida tiene 4 hojas:
- **Dashboard** — resumen del proceso
- **Layout CONTPAq** — pólizas listas para importar (bloques P / M1 / AD)
- **Resumen Matches** — detalle de cada match banco ↔ factura
- **Sin Match - Revisar** — movimientos que requieren revisión manual

---

## Motor de matching (3 capas)

El sistema intenta identificar automáticamente cada depósito:

1. **Folio + RFC** — si el concepto bancario contiene el número de folio y coincide con el RFC del pagador
2. **Monto exacto + RFC** — si el monto del depósito coincide con el total de la factura (±$0.10)
3. **RFC único** — si el RFC del pagador tiene solo una factura pendiente

Los movimientos que no pasan ninguna capa quedan en la hoja "Sin Match - Revisar".

---

## Estructura del layout CONTPAq

```
P  | Fecha        | 1 | num_poliza | 1 | 0 | COBRANZA CLIENTE | 11 | 0 | 0
M1 | 10201001     | Folio | 0 | Total c/IVA | 0 | 0 | COBRANZA CLIENTE
M1 | 20901000     | Folio | 0 | IVA         | 0 | 0 | COBRANZA CLIENTE
M1 | 20801000     | Folio | 1 | IVA         | 0 | 0 | COBRANZA CLIENTE
M1 | 1050XXXX     | Folio | 1 | Total c/IVA | 0 | 0 | COBRANZA CLIENTE
AD | UUID-factura
```

---

## Actualización del catálogo de cuentas

Cuando se agrega un cliente nuevo:
1. Abre el Excel de pólizas del mes anterior
2. Ve a la hoja `cuentas`
3. Agrega una fila con: `C | 1050XXXX | Nombre del cliente | 1050XXXX`
4. Sube ese archivo actualizado en el campo "Catálogo" la próxima vez

---

## Soporte
GBC Business Consulting — Uriel / Jessica / Roberto
