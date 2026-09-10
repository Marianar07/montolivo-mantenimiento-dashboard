# Cómo actualizar el tablero de mantenimiento

Guía rápida para cuando tengas datos nuevos del CMMS.

## Cada vez que quieras actualizar el tablero

1. Exporta los 4 archivos de siempre desde el CMMS (Órdenes de Trabajo,
   Solicitudes de Servicio, Disponibilidad, Datos Generales de Equipos —
   este último trae criticidad y "Provoca Paro?" de cada activo) y
   guárdalos dentro de la carpeta `datos_nuevos` (reemplazando los que
   estaban ahí de la vez anterior).

2. Abre una terminal en esta carpeta (`Mantenimiento`) — en el Explorador de
   Windows, clic derecho dentro de la carpeta → "Abrir en Terminal" (o
   "Abrir ventana de PowerShell aquí").

3. Escribe:
   ```
   claude
   ```

4. Cuando se abra Claude Code, escríbele algo como:
   > Actualiza el tablero con los archivos que puse en datos_nuevos y sube
   > los cambios a GitHub

5. Claude Code va a procesar los archivos, generar el tablero nuevo, y
   subirlo a tu repositorio. Cuando te confirme que terminó, revisa el
   enlace de Vercel — en uno o dos minutos ya debería mostrar los datos
   actualizados.

## Si algo no cuadra

Si Claude Code te avisa que algún número raro apareció (por ejemplo, muy
pocas OT quedaron ligadas a un equipo, o los totales de criticidad
cambiaron mucho), probablemente el CMMS cambió algo en el formato del
export. En ese caso, es mejor traer los archivos a una conversación conmigo
(Claude, aquí en Cowork) para revisar juntos qué cambió antes de publicar el
tablero.

## Qué NO tocar

- No edites `index.html` ni `data.json` a mano — se regeneran solos cada
  vez que corres el proceso.
- Los umbrales de semáforo (colores verde/amarillo/rojo) y la lógica de
  cálculo están documentados en `CLAUDE.md`, por si algún día el jefe de
  mantenimiento pide cambiarlos.
