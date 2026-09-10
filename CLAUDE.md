# Tablero de Mantenimiento — Inversiones Montolivo

Este proyecto genera un tablero de indicadores de mantenimiento (`index.html`,
un solo archivo autocontenido) a partir de 4 exports en Excel del CMMS. El
tablero se despliega en Vercel conectado al repositorio de GitHub de este
proyecto: cada vez que se sube un `index.html` nuevo al repo, Vercel lo
redespliega automáticamente en el mismo enlace.

## Qué hacer cuando Mariana pida "actualizar el tablero"

1. Verificar que los 4 archivos Excel más recientes estén en `datos_nuevos/`
   (si no están, pedirle que los coloque ahí antes de continuar). Los 4 tipos
   de archivo, por su forma típica de nombre en el export del CMMS:
   - **OT (Órdenes de Trabajo)**: nombre tipo `Informe_1.xlsx` — hoja única,
     columnas incluyen `Código O.T.`, `Entidades`, `Tipo`, `Prioridad`,
     `Fecha Creación`, `Fecha Inicio Programado/Real`, `Fecha Fin
     Programado/Real`, `Estado O.T.`, `Ejecución (%)`, `Ejecutores`,
     `Descripción`, `Total Real`.
   - **SS (Solicitudes de Servicio)**: nombre tipo `Informe.xlsx` — columnas
     incluyen `Código`, `Entidad`, `Tipo`, `Prioridad`, `Estado`, `Fecha de
     solicitud`, `Fecha de respuesta`, `OTs`.
   - **Disponibilidad**: nombre contiene "Disponibilidad" — hoja
     `Disponibilidad`, la fila 1 es un encabezado de periodo ("Periodo:
     Desde ... Hasta ..."), los nombres de columna reales están en la fila 2.
     Columnas: `Código`, `Equipo`, `Instalación de Proceso`, `Tiempo
     Producción Hábil/Real [Horas]`, `Tiempo Paro Correctivo/Preventivo
     [Horas]`, `Disponibilidad [%]`.
   - **Datos Generales de Equipos**: nombre contiene "Datos Generales de
     Equipos" — hoja única `Sheet`, columnas incluyen `Código`, `Nombre`,
     `Criticidad` (Alta/Media/Baja) y `Provoca Paro?` (Sí/No) para cada uno
     de los ~2323 activos del maestro. Reemplaza al antiguo export "EQxIP"
     (hoja `Novedades`) — ver nota en "Paros por equipo" más abajo.

2. Ejecutar el script de procesamiento:
   ```
   python process_dashboard.py
   ```
   Esto lee los 4 Excel de `datos_nuevos/` (o de la carpeta raíz del
   proyecto si `datos_nuevos/` está vacía), regenera `data.json` e inyecta
   ese JSON en `index_template.html` para producir el `index.html` final.

   Si `pandas`/`numpy`/`openpyxl` no están instalados, correr primero:
   ```
   pip install -r requirements.txt
   ```

3. **Revisar el reporte que imprime el script en consola** antes de subir
   nada: cuántas OT se resolvieron a un equipo/lugar, cuántos paros se
   identificaron, los totales de criticidad. Si algún número cambia de forma
   extraña respecto a la corrida anterior (por ejemplo, el % de OT resueltas
   cae mucho), es señal de que el CMMS cambió el formato de exportación o
   aparecieron ubicaciones/equipos nuevos que la lógica de resolución no
   contempla — avisar a Mariana en vez de subir el archivo directamente.

4. Confirmar visualmente que el `index.html` generado se ve bien (abrirlo en
   el navegador, o usar Playwright si está disponible) — sobre todo la
   pestaña "Resumen Gerencial" y "Paros por Equipo".

5. Subir el cambio al repositorio de GitHub:
   ```
   git add index.html data.json
   git commit -m "Actualización de datos: <fecha>"
   git push
   ```
   Vercel detecta el push y redespliega solo, sin pasos adicionales.

6. Confirmarle a Mariana qué se actualizó (fecha de los datos, cambios
   relevantes en los indicadores) y que el enlace de Vercel ya refleja los
   datos nuevos.

## Lógica de negocio importante (no romper al modificar el script)

### Resolución de equipo y lugar (OT/SS -> Disponibilidad)

El campo `Entidades` (OT) / `Entidad` (SS) tiene el formato
`"CÓDIGO | NOMBRE"`. Para ubicar el equipo y el lugar:

- Si `CÓDIGO` existe directamente como `Código` en el archivo de
  Disponibilidad → es un equipo específico (`equipo_especifico`); el lugar
  sale de su columna `Instalación de Proceso` (parte después del `|`).
- Si `CÓDIGO` empieza con `MTV` (es un código de ubicación, no de equipo
  puntual) → buscar en Disponibilidad un activo cuyo código empiece con
  `AI-` y cuya `Instalación de Proceso` sea exactamente igual al texto
  completo de `Entidades`/`Entidad` (estos activos "AI-xxxx" tienen nombre
  `"ADECUACIÓN E INSTALACIÓN <lugar>"` y sirven de sustituto cuando la OT/SS
  solo referencia el punto de venta, no un equipo concreto). Este caso se
  marca `instalacion_pdv`.
- Si no se encuentra ninguno de los dos casos, queda `solo_ubicacion` (se
  conoce el lugar pero no un activo) o `sin_dato` (nada resuelto).

Esta lógica surgió de una corrección explícita de Mariana: al principio solo
se resolvían 6 de 53 OT correctivas por código directo; con el patrón
`AI-xxxx` se llegó a 47/53 (89%). **No volver al enfoque de solo-código
directo** — se pierde la mayoría de la cobertura.

### Paros por equipo

Un "paro" es una OT de tipo Correctivo cuyo equipo (`equipo_cod`, ya resuelto
con la lógica de la sección anterior) está marcado `Provoca Paro?` = `"Sí"`
en **`Datos Generales de Equipos.xlsx`** (el maestro de activos que también
trae la `Criticidad`, ver arriba). Es la clasificación oficial del CMMS
sobre qué activos, al fallar, generan un paro real de operación —
`construir_paros()` arma este set directamente desde el dataframe de
criticidad ya cargado (`crit`), sin necesitar un archivo aparte.

**Historial:** antes de tener la columna `Provoca Paro?`, el criterio era
una aproximación: contaba como paro cualquier correctivo cuyo equipo/lugar
se hubiera podido resolver (`equipo_especifico` o `instalacion_pdv`, ver
sección anterior) — usando de forma indirecta la marca de activos "no
mantenibles" ANM y el prefijo `AI-` de los activos sustitutos de punto de
venta. Ese proxy sobrestimaba los paros: los activos `AI-xxxx |
ADECUACIÓN E INSTALACIÓN <lugar>` (los sustitutos que se usan solo para
poder ubicar equipo/lugar cuando la OT referencia el punto de venta, no un
equipo puntual — ver sección anterior) casi todos están marcados `"No"` en
`Provoca Paro?`, porque no son equipos reales. **No volver a ese criterio
(ANM/`AI-`) para decidir qué es un paro** — usar siempre `Provoca Paro?`.
(La marca ANM se sigue usando, sin relación con esto, para excluir activos
del directorio de equipos en la pestaña "Equipos" — ver más abajo.)

El sistema CMMS todavía no tiene un módulo de paros propio con datos
reales — este indicador sigue siendo un cálculo derivado, no un reporte
nativo. Cada paro puede estar:
- **Pendiente de iniciar**: la OT sigue abierta y no tiene `Fecha Inicio
  Real` — se usa `Fecha Inicio Programado` como fecha de referencia, y la
  duración queda `null`.
- **En curso**: tiene `Fecha Inicio Real` pero no `Fecha Fin Real` —
  duración `null`.
- **Finalizado**: tiene ambas fechas — duración = diferencia en horas.

### Enlace SS -> OT

El campo `Descripción` de la OT a veces contiene una referencia de texto
libre tipo `"SS-00004"`. Se extrae con la expresión regular `SS-0*(\d+)` y
se cruza contra la columna `Código` del archivo de SS (rellenando a 5
dígitos con ceros a la izquierda). El archivo de SS puede traer códigos
duplicados exactos (mismo número, fila repetida) — el script los
deduplica quedándose con la primera ocurrencia.

### Semántica de "Fecha de respuesta" en SS (confirmado con Mariana)

El campo `Fecha de respuesta` del archivo de SS **marca el cierre de la
solicitud, no una simple revisión**. Confirmado cruzando Estado vs. Fecha de
respuesta en los datos reales:
- **Creada, Leída, Programada en O.T., Validada** → 0% tienen Fecha de
  respuesta llena. Son estados de una SS que **sigue abierta**, sin solución
  todavía.
- **Evaluada** (mantenimiento programado, casi siempre) y **No aprobada**
  (la solicitud es duplicada, se creó mal, o ya se solucionó de otra forma)
  → 100% tienen Fecha de respuesta llena, y se cierran el mismo día de esa
  fecha.

Por lo tanto: "Fecha de respuesta" menos "Fecha de solicitud" **sí es un
tiempo de solución válido**, calculado solo sobre las SS que ya están
cerradas (Evaluada o No aprobada) — no es un promedio de "primera revisión".
El tablero debe dejar esto explícito en la etiqueta/nota de esa métrica
(algo como "Tiempo promedio de solución (SS cerradas)"), y complementarlo
con:
1. Un listado de "Novedades abiertas" (SS sin Fecha de respuesta: Creada,
   Leída, Programada en O.T., Validada), ordenado de más antigua a más
   reciente por Fecha de solicitud, con código, tipo, severidad, equipo,
   lugar, fecha de solicitud y días que lleva abierta.
2. Un desglose del tiempo promedio de solución por Tipo de SS y por
   Severidad/Prioridad (igual que ya existe para las OT), calculado solo
   sobre las SS cerradas.

Ojo: como normalmente hay muy pocas SS cerradas en un momento dado (10 de
102 en la corrida de referencia), estos promedios son sobre una muestra
chica — no ocultarlo, mostrar siempre el "N sobre el cual se calculó".

### Equipos "no mantenibles" (ANM)

En el archivo de Criticidad/Disponibilidad, los activos cuya `Instalación`
empieza con `"ANM |"` (Activos No Mantenibles) se excluyen del directorio
de equipos del tablero — no son equipos operativos de mantenimiento.

### Limitaciones de datos conocidas (comunicadas a Mariana, no ocultarlas)

- El módulo de Disponibilidad del CMMS todavía muestra 100% de
  disponibilidad y 0 horas de paro para **todos** los equipos — no refleja
  los paros reales todavía. El tablero lo señala con un aviso visible en la
  pestaña "Disponibilidad y Criticidad".
- El cumplimiento preventivo y el OTIF pueden verse artificialmente altos
  porque el CMMS a veces registra las fechas de cierre/programación de
  forma retroactiva (se iguala la fecha programada a la fecha real).
- El MTBF usa un periodo de referencia fijo (3360 h), no horas de operación
  específicas por equipo — es un valor referencial.
- "Confiabilidad de inventarios" y "Eficiencia de personal" no son
  calculables con los archivos actuales — no existen en el tablero.

## Umbrales de semáforo (provisionales)

Definidos en `process_dashboard.py`, diccionario `UMBRALES` — pendientes de
validar con el jefe de mantenimiento:
- Disponibilidad: verde ≥95%, amarillo 90–95%, rojo <90%
- Cumplimiento preventivo: verde ≥90%, amarillo 75–90%, rojo <75%
- OTIF: verde ≥90%, amarillo 75–90%, rojo <75%
- % paradas no programadas: verde <5%, amarillo 5–10%, rojo >10%

## Estructura del proyecto

```
Mantenimiento/
├── CLAUDE.md              <- este archivo (instrucciones para Claude Code)
├── LEEME.md                <- guía rápida en lenguaje simple para Mariana
├── process_dashboard.py    <- script de procesamiento (Python)
├── requirements.txt         <- dependencias del script
├── index_template.html      <- plantilla del tablero (CSS/JS), con marcador __DATA_JSON__
├── index.html                <- tablero final generado (el que se sube a GitHub)
├── data.json                 <- datos procesados más recientes (se regenera cada vez)
└── datos_nuevos/              <- Mariana coloca aquí los 4 Excel de cada actualización
```

**No editar `index.html` ni `data.json` a mano** — siempre regenerarlos
corriendo `process_dashboard.py`. Si hay que cambiar el diseño, los
colores, o agregar una pestaña, el cambio va en `index_template.html`
(HTML/CSS/JS) y luego se vuelve a inyectar el JSON.

## Identidad de marca

Colores extraídos del tablero de referencia de la empresa (paleta ya
aplicada en `index_template.html`, no recalcular):
```
--maroon-dark:#52101E;  --maroon:#7A1A2E;  --cream:#F8F3EE;
--good:#16A34A; --warn:#D97706; --bad:#DC2626;
--chart-1:#2a78d6; --chart-2:#eb6834; --chart-3:#1baf7a;
```
