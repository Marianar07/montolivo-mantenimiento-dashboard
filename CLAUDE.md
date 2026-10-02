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

   También opcional pero recomendado: **`Datos Generales Activos.xlsx`**
   (nombre contiene "datos generales activos"; agregado por Mariana el
   1-oct-2026 para tener los equipos nuevos). Es el export de **Activos**
   del CMMS, más actualizado que Disponibilidad, con otro formato: columnas
   `Codigo` (sin tilde), `Nombre`, `Código IP.`, `Instalación Proceso`,
   `Provoca Paro` (sin "?") y **sin `Criticidad`**. `integrar_activos()` lo
   usa como maestro de activos: agrega a Disponibilidad los activos que
   faltan (salen en el directorio y en el total de equipos, sin % de
   disponibilidad → "—"), actualiza la ubicación de los trasladados, toma
   `Provoca Paro` de aquí y conserva la `Criticidad` de "Datos Generales de
   Equipos" (los nuevos quedan sin criticidad). El script imprime cuántos
   activos nuevos, trasladados y sin criticidad hay. 1-oct-2026: 2393 activos
   (70 nuevos: 68 granizadoras 2T + nevera Coca-Cola y granizadora Artiq de
   Arkadia, estas dos "Provoca Paro" = Sí), 15 trasladados.

   Además, opcionalmente, **`TECNICOS.xlsx`** (nombre contiene "tecnicos")
   — maestro de técnicos de planta, columnas `NOMBRE`, `CEDULA`, `CIUDAD`.
   No es uno de los 4 exports periódicos del CMMS — es una lista de
   personal que cambia poco, así que no hace falta pedirlo cada vez; solo
   volver a colocarlo si hay altas/bajas de técnicos. Si no está, el
   tablero simplemente no separa "trabajo de terceros" (ver más abajo).

   También opcional: **`Datos Generales de Proveedores.xlsx`** (nombre
   contiene "proveedores") — maestro de proveedores del CMMS, hoja única
   `Sheet`, columna `Nombre`. Igual que `TECNICOS.xlsx`, cambia poco y no
   hace falta pedirlo cada vez. Un ejecutor de OT se clasifica en tres
   grupos (ver detalle en "Técnicos internos vs. proveedores vs. otros" más
   abajo): **técnico interno** (está en `TECNICOS.xlsx`/`ALIAS_TECNICOS`),
   **proveedor** (no es interno, pero su nombre aparece tal cual en este
   maestro de proveedores o en `PROVEEDORES_EXTRA` del script — p. ej.
   "SEBASTIAN BUITRAGO GRACIANO", que ejecuta a nombre de un proveedor), u
   **"otro"** (no es interno y tampoco aparece en ninguno de los dos). Si
   no se encuentra el archivo, ningún ejecutor no interno se reconoce como
   proveedor (quedan como "otro"; el tablero igual los lista con los
   proveedores, marcados "No registrado").

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
nativo.

**Paros correctivos arrancan en la SS** (definición de Mariana, 28-sep-2026,
`construir_paros_correctivos()`): un paro empieza cuando se reporta una SS
sobre un equipo con `Provoca Paro?` = "Sí" (código de la `Entidad` de la SS).
Estados:
- **Fuera de servicio**: reportado, nadie ha empezado a trabajar (la SS no
  tiene OT todavía, o la OT no tiene `Fecha Inicio Real`).
- **En reparación**: la OT tiene `Fecha Inicio Real` y no `Fecha Fin Real`.
- **Finalizado**: la OT tiene `Fecha Fin Real` (o la SS se cerró sin OT →
  su `Fecha de respuesta`). Duración = fin − inicio del paro.

Agrupación: SS enlazadas a la misma OT = un paro (arranca en la SS más
antigua); SS abiertas sin OT del mismo equipo = un paro; SS "No aprobada" no
son paro; OT correctivas sobre equipos que generan paro sin ninguna SS
también son paro (arrancan en la `Fecha Creación` de la OT). El inicio nunca
es posterior a la `Fecha Inicio Real` (hay OT registradas después de hacer
el trabajo). El enlace SS → OT usa la columna `OTs` del export de SS (trae
el **código** de la OT, verificado 92/92) y, de respaldo, la referencia
"SS-xxxxx" en la Descripción de la OT.

**Preventivos** (`paros_programados`, `construir_paros(..., tipo="Preventivo",
solo_iniciadas=True)`): solo cuentan desde que el técnico inicia — estados
"En mantenimiento" → "Finalizado". El estado lo calcula el script en
`estado_paro`; el tablero lo pinta con `badgeEstadoParo()`.

En el Resumen, la tarjeta "Paros en el periodo" tiene tres filas
(pedido de Mariana): **Correctivos** y **Preventivos** (solo cantidad y
duración promedio, sin clic) y **Sin resolver** (la única clicable → lista) (sin fin, resaltada en rojo; su duración se muestra como
"N días y contando" hasta hoy). En la esquina: duración promedio de todos los
paros y cada fila su promedio; un paro cuenta con su `duracion_h` si ya
terminó o con lo que lleva parado **hasta hoy** si sigue sin resolver (`durParoH`).
**No volver a promediar solo los finalizados**: excluía los sin resolver,
que son los más largos, y el promedio salía muy bajo (1,1 días vs. 5,2 reales
el 29-sep-2026). Menos de un día se
muestra en horas, si no en días.

La pestaña "Paros por Equipo" **respeta el filtro de fechas y de lugar**
(pedido de Mariana; antes lo ignoraba y mostraba solo los paros abiertos):
muestra los paros (correctivos + preventivos) cuya fecha de paro cae en el
rango, finalizados incluidos — mismo criterio y mismo total que "Paros en el
periodo" del Resumen Gerencial. "Frecuencia de mantenimiento por equipo"
usa las OT creadas en el rango.

### Una fila por OT (el export trae una fila por actividad)

El export de OT trae **una fila por cada actividad** de la OT (columna
`Actividades`), cada una con su propio `Total Real`; el resto de campos se
repite. `construir_ot()` consolida a una fila por `Código O.T.` sumando el
`Total Real` de todas sus filas. Confirmado por Mariana contra Mantum: OT
000022 = $128.242,85 y OT 000182 = $243.749,90 (= suma de sus filas, aunque
en la 000182 las 5 filas traen el mismo valor). No contar filas como OT ni
tomar un solo costo por OT.

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
respuesta en los datos reales (corrida de referencia inicial y re-validado
el 23 de septiembre de 2026 con datos más recientes, que ya incluían los
estados nuevos Ejecutada/Editada):
- **Creada, Editada, Leída, Programada en O.T., Validada** → 0% tienen
  Fecha de respuesta llena. Son estados de una SS que **sigue abierta**,
  sin solución todavía.
- **Ejecutada** (se resolvió con una OT), **Evaluada** (mantenimiento
  programado, casi siempre) y **No aprobada** (la solicitud es duplicada,
  se creó mal, o ya se solucionó de otra forma) → 100% tienen Fecha de
  respuesta llena, y se cierran el mismo día de esa fecha.

El CMMS ha ido agregando estados nuevos con el tiempo (p. ej. Ejecutada y
Editada no existían en la corrida de referencia original) — **por eso el
tablero (`index_template.html`) nunca debe decidir "abierta vs. cerrada"
comparando el texto del Estado contra una lista fija de valores; siempre
debe usar la presencia de `Fecha de respuesta` como criterio** (ya está
así implementado — ver `getFilteredSS`/`ssAbiertas` en Resumen Gerencial y
`conRespuesta` en la pestaña Solicitudes de Servicio). Si aparece un estado
nuevo en una corrida futura, revisar con el mismo cruce Estado vs. Fecha de
respuesta si debe listarse como abierto o cerrado en las notas de la UI,
pero no hace falta tocar la lógica de cálculo.

En la pestaña **Solicitudes de Servicio**, las tarjetas **SS abiertas**
(Creada, Editada, Leída, Validada, Programada en O.T.) y **SS cerradas**
(Ejecutada, Evaluada, No aprobada), cada una con el conteo por estado,
reemplazaron a "Pendientes de gestión (Creada)" (pedido de Mariana,
1-oct-2026: las pendientes no son solo las Creada). Se calculan por
`Fecha de respuesta`, igual que en el Resumen.

Por lo tanto: "Fecha de respuesta" menos "Fecha de solicitud" **sí es un
tiempo de solución válido**, calculado solo sobre las SS que ya están
cerradas (Ejecutada, Evaluada o No aprobada) — no es un promedio de
"primera revisión". El tablero debe dejar esto explícito en la
etiqueta/nota de esa métrica (algo como "Tiempo promedio de solución (SS
cerradas)"), y complementarlo con:
1. Un listado de "Novedades abiertas" (SS sin Fecha de respuesta),
   ordenado de más antigua a más reciente por Fecha de solicitud, con
   código, tipo, severidad, equipo, lugar, fecha de solicitud y días que
   lleva abierta.
2. Un desglose del tiempo promedio de solución por Tipo de SS y por
   Severidad/Prioridad (igual que ya existe para las OT), calculado solo
   sobre las SS cerradas.

Ojo: como normalmente hay muy pocas SS cerradas en un momento dado (10 de
102 en la corrida de referencia original), estos promedios son sobre una
muestra chica — no ocultarlo, mostrar siempre el "N sobre el cual se
calculó". (El cuadro "Tiempo de solución promedio" se muestra en días y sin
texto gris debajo, a pedido de Mariana; el N queda visible en la nota azul
de arriba de la pestaña: "N de M SS en el rango".)

### Bloque "Órdenes de trabajo" del Resumen Gerencial

Pedido de Mariana (29-sep-2026), en `renderResumen()`: tres tarjetas —
**OT totales**, **OT abiertas** y **OT cerradas**, cada una con el conteo
por tipo debajo del número (Correctivo · Preventivo · cualquier otro tipo del
CMMS, p. ej. "Correctiva programada", para que sume el total). Ya no hay tarjeta aparte
"OT para reprogramar": dentro de "OT abiertas" hay una fila **"Vencidas"**
clicable → listado con días de atraso. **Vencida = para reprogramar**
(definición de Mariana): OT abierta cuya `Fecha Fin Programado` ya pasó y
que no está hecha — sin `Fecha Fin Real` **o con `Ejecución (%)` < 100**
(`necesitaReprogramar()`; días de atraso desde la fin programada). Las
abiertas con Fecha Fin Real y 100% de ejecución (trabajo hecho, falta
cerrarlas en el CMMS, `otEjecutadaCompleta()`) van aparte en la fila
"Ejecutadas sin cerrar" (`sinCerrarEnSistema()`), también clicable (listado
con inicio real, fin real y % de ejecución; el de vencidas trae inicio y
fin programado y % de ejecución). La
sección "Órdenes de Trabajo para reprogramar" de la pestaña OT usa el
mismo criterio. "OT totales" muestra debajo el
cumplimiento preventivo y correctivo (mini barras; ya no hay tarjetas
individuales de cumplimiento en "Lo más importante"). Abierta/cerrada según el `Estado
O.T.` (`estado==='Abierta'`).
- Cerradas: tiempo promedio de solución = `Fecha Fin Real` − `Fecha
  Creación` (en días, solo las que traen Fecha Fin Real; se muestra el N),
  desglosado por severidad.
- Abiertas: días promedio que llevan abiertas = hoy − `Fecha Creación`,
  desglosado por severidad (el promedio general va en la esquina superior
  derecha de la tarjeta, igual que en OT cerradas y en las SS); cada severidad es clicable y abre el listado de
  esas OT (con días abierta; sin columna de "vencida", a pedido de Mariana).

- **"Creadas después del inicio"** (al final de OT totales, resaltada en
  rojo con `.sev-row-alerta`, pedido de Mariana): OT
  cuya `Fecha Creación` es posterior a su `Fecha Inicio Real` — registradas
  en Mantum cuando el trabajo ya había empezado. Se muestra "N de M" (M =
  OT con inicio real); clic → listado solo con OT, entidad, lugar,
  técnico, fecha de creación, fecha de inicio (real) y cuánto después se
  registró (sin tipo, estado ni costo, a pedido de Mariana). (Con inicio *programado* en vez de real serían más:
  162 de 244 el 29-sep-2026.)
- **OT de proveedores**: OT con algún ejecutor proveedor (`esOTProveedor`,
  incluye `PROVEEDORES_EXTRA`), con tipo y filas **Abiertas** / **Cerradas**
  clicables → listado con columna Proveedor (Mariana quitó las filas por
  nombre de proveedor).
- **OT sin técnico asignado**: OT con `Ejecutores` vacío; clic → listado.
- **Costos** (pedido de Mariana): las tarjetas de OT **no** muestran costo
  (se quitó); el costo va solo en los listados: todos los de este bloque
  (`modalOT()`) llevan columna Costo y el costo total de lo filtrado. Las
  severidades de "OT cerradas" también son clicables (listado con fin real,
  días de solución y costo).
- **Filtros en los listados**: todos los listados emergentes del Resumen
  (OT y "Paros en el periodo") usan `modalFiltrable()`: buscador de texto
  (sin tildes, sobre todas las columnas) + desplegables de Tipo, Estado,
  Severidad y Lugar (solo los que tengan más de un valor); muestra "N de M"
  y recalcula el costo total con lo filtrado.
- **Columna Comentarios** (pedido de Mariana, 2-oct-2026): `modalFiltrable()`
  la agrega sola al final de todo listado cuyas filas traen `comentarios`
  (OT y SS). OT = `Realimentación` (sin el bloque "Ejecutores:");
  SS = columna `Comentarios` del export (historial "Estado / Persona /
  fecha / texto", más reciente primero, un renglón por cambio;
  `limpiar_comentarios_ss()` parte por "SEPARADOR-COMENTARIOS").
- **Costo total mantenimiento** es clicable → tabla de costo por lugar (OT,
  correctivo, preventivo, correctiva programada, total y % del total), con
  buscador; respeta el filtro de fechas/lugar como el resto del Resumen.
- **Desglose del costo** ("Lo más importante", junto a "Costo total"): el
  export de OT solo trae `Total Real` por OT. **Costo servicios
  (proveedores)** = Total Real de las OT de proveedores (clicable).
  **Costo mano de obra** = Total Real de las OT **sin proveedores**
  (`otManoObra`; clicable → costo por lugar de esas OT). Mariana, 1-oct-2026:
  no puede incluir a los terceros porque ya están en "Costo servicios", así
  que **Mano de obra + Servicios = Costo total**. **Mariana confirmó el 1-oct-2026**:
  el costo total es costo de mano de obra (`Total Real` = horas reales × una
  tarifa fija por técnico, p. ej. $24.427/h Alejandro, $17.009/h Juan
  José/Jonathan). Antes (29-sep) la tarjeta decía "Sin datos" porque no
  estaba confirmado. **Costo repuestos** sigue "Sin datos" (Mantum no lo
  reporta). Validado el 1-oct-2026 con `Datos Generales de Recursos.xlsx`
  (catálogo de recursos asignables a una OT, cargado desde SIESA): 205
  repuestos, **todos con `Valor Unitario` = 0**. El export de OT ya trae
  columnas `Código recurso`, `Nombre recurso`, `Cantidad real`, `Cantidad
  estimada`, pero solo 1 OT (000082, RP0130 ×1) tiene un recurso asignado.
  Cuando los recursos traigan precio y se asignen en las OT, costo
  repuestos = Σ `Cantidad real` × `Valor Unitario` (cruce por código). Costo servicios (proveedores) sale del mismo Total Real, por eso
  se resta de mano de obra (no se cuenta dos veces).

### Bloque "Solicitudes de servicio" del Resumen Gerencial

Pedido de Mariana (29-sep-2026), mismo estilo que el bloque de OT:
- **SS totales**: abiertas/cerradas (por `Fecha de respuesta`) y
  cumplimiento vs. Fecha esperada (`cumplimientoStatsSS()`, el mismo de la
  pestaña SS: cumplidas / incumplidas / pendientes). Clic → todas las SS.
- **SS abiertas** y **SS cerradas** (Mariana, 29-sep-2026: se quitaron las
  tarjetas por estado). Cerradas = Ejecutada, Evaluada y No aprobada; el
  resto abiertas — el código lo decide por `Fecha de respuesta`, que coincide
  100% con esa lista y clasifica solo cualquier estado nuevo. Cada tarjeta:
  días promedio abiertas (o de solución), filas por severidad (con sus días,
  clicables → listado) y, debajo del número, el conteo por estado en
  línea ("Creada: 135 · Editada: 66…", sin bolitas de color — Mariana las
  quitó por saturar) — sin clic
  (el estado se filtra dentro del listado). Los días promedio (abiertas / solución) van
  en la esquina superior derecha de la tarjeta (`kpiCard({corner})`).
- Los listados usan `modalFiltrable()` (buscador + filtros de Tipo, Estado,
  Severidad, Lugar) con columnas SS, tipo, severidad, estado, equipo, lugar,
  solicitud, esperada, respuesta, días, cumplimiento y OT. Esperada y
  respuesta muestran también la **hora** (`fmtFechaHoraCorta`): el
  cumplimiento compara fecha y hora, y sin la hora una SS respondida el mismo
  día de la esperada pero más tarde parecía mal marcada como incumplida.
  En los listados de SS abiertas (ninguna con respuesta) no se muestra la
  columna Respuesta.
Reemplazó a la tarjeta suelta "Solicitudes de servicio" del grupo "Servicio".
- **SS con OT asignada** (pedido de Mariana, 1-oct-2026, `tarjetaSSOT`):
  SS con OT en la columna `OTs` del export (`ot_asociada`), con filas **OT
  abierta** / **OT cerrada** según el `Estado O.T.` de esa OT (buscada en
  todas las OT, no solo las del filtro) y aparte **SS sin OT asignada**;
  cada fila con su % sobre las SS del rango y clicable → `modalSS`. Si
  alguna OT enlazada no aparece en el export sale una fila "OT no
  encontrada" (hoy 0; 1-oct-2026: 77 de 397 con OT, 33 abierta / 44 cerrada).

**Códigos de SS**: en todo el tablero se muestran solo con el número
(p. ej. `00016`), sin prefijo "SS-" (pedido de Mariana; ver `fmtSSParo()`).

### Horas de técnicos en el Resumen

La tarjeta **"Técnicos"** (antes "Técnicos en sitio / programados") muestra
la cantidad de técnicos internos activos y, debajo, cuántos están En sitio /
Programados / Sin asignación (mismas categorías de "Ubicación del técnico",
`calcularUbicacionTecnicos`). Además (pedido de Mariana): **Tiempo programado (h)** (antes "Horas legales") = horas hábiles del rango (42 h/semana,
`horasDisponiblesRango`) × técnicos considerados (internos activos + los
inactivos que trabajaron en el rango); **Tiempo trabajado (h)** = suma de la
duración real de las OT, sin tope; **Ocupación** = trabajado / programado. Sale de
`calcularCapacidadTecnicos()`, la misma función de la pestaña Técnicos, así que
los números coinciden con "Capacidad y ocupación por técnico". Los técnicos de
TECNICOS.xlsx sin ninguna OT (hoy Diego Ocampo, Jhonier Meneses, Pedro Lozada)
cuentan en horas legales con 0 trabajadas, y bajan la ocupación.
Clic en la tarjeta → "Tiempos por técnico" (incluye inactivos que trabajaron
en el rango, marcados "Inactivo"): cada técnico con estado ahora
(en sitio/programado), OT con duración, tiempo programado, tiempo trabajado,
libres y % ocupación.

### "OT por tipo de mantenimiento — histórico por mes" (Resumen Gerencial)

Pedido de Mariana (reemplazó a "OT por tipo y severidad" y luego al
histórico tipo × severidad): `chartTipoMensual()` — barras por mes (mes de
Fecha Creación) **desde agosto 2026** (`HISTORICO_DESDE`; antes solo hay una
OT suelta de mayo), una barra por tipo presente (Correctivo vino, Preventivo
bronce, otros tipos como Correctiva programada en pizarra), con el número
encima (`groupedBarChart` con `fill` y `showValues`). Incluye meses vacíos
del medio. Respeta el filtro general.

### "OT correctivas por lugar" (Resumen Gerencial)

Reemplazó a "Top lugares con más averías correctivas" (pedido de Mariana):
muestra **todos** los lugares (no solo 8) con la cantidad de OT correctivas
del rango del filtro general de arriba (ya **no** tiene selector de mes propio
ni el total de correctivas, a pedido de Mariana); la cantidad es clicable → listado de esas OT con filtros y costo (`topLugaresAverias()`,
`renderTopLugaresWrap()`).

### "OT por lugar y tipo de mantenimiento" (Resumen Gerencial)

Pedido de Mariana (1-oct-2026), tarjeta a todo el ancho debajo de las dos
anteriores (`chartTipoLugar()`): una barra horizontal por lugar, segmentada
por tipo (mismos colores que el histórico mensual), ordenada por total, con
la cantidad en cada segmento y el total al final. Respeta el filtro general.
Clic en el nombre del lugar → listado de sus OT (`modalFiltrable`, con filtro
de Tipo y costo). Las OT con varias entidades traen el lugar como
"SAN DIEGO, SJ-0016 | ..." — `lugarCortoOT()` las agrupa en su lugar.

### Gráficas de SS (Resumen Gerencial)

Pedido de Mariana (1-oct-2026), al final del Resumen, lado a lado:
**"SS por tipo — histórico por mes"** (`chartTipoMensual(ss, 'fecha_ss')`,
desde agosto 2026 como el de OT) y **"SS por punto de venta"**
(`chartTipoLugar(ss, 'data-ss-lugar')`, barras por lugar segmentadas por
tipo de SS; clic en el lugar → `modalSS`). Ambas reutilizan las funciones
de las gráficas de OT; los colores por tipo salen de `tiposYColores()`
(Correctivo vino, Preventivo bronce, el resto por frecuencia). Los tipos de
SS se muestran tal cual vienen del CMMS (p. ej. "Correctivos de emergencia"
y "Correctiva de Emergencia" quedan separados).

### Ordenar tablas por columna

Pedido de Mariana: **todas** las tablas se ordenan al hacer clic en el
título. Las paginadas (`buildPagedTable`) ordenan sus datos completos
(`.sortable-th`, `sortVal`); todas las demás (listados emergentes, tablas
pequeñas) usan el manejador global `ordenarTablaPorColumna()`/`valorOrden()`,
que ordena las filas visibles por el texto de la celda entendiendo dinero,
números, %, horas/días y fechas dd/mm/aaaa. Una tabla nueva queda ordenable
sola; para excluir una, envolverla en un elemento con `data-no-orden`.

### Promedios por día: SS creadas y OT realizadas

Dos cuadros usan la misma función `promedioPorDia()`:
- "SS creadas por día (promedio)" (pestaña Solicitudes de Servicio): SS
  por `fecha_ss`.
- "OT realizadas por día (promedio)" (pestaña Órdenes de Trabajo): OT con
  `Fecha Fin Real`, contadas por esa fecha (no por la de creación).
- "OT creadas por día (promedio)" (pestaña Órdenes de Trabajo, pedido de
  Mariana 1-oct-2026, al lado del anterior): OT por `Fecha Creación`.

Cálculo = registros en el rango / días calendario del rango (incluye
domingos). El rango se acota a las fechas que traen los datos (un filtro que
llega al futuro no diluye el promedio). Si no hay "Desde" elegido, arranca
en `INICIO_USO_CMMS` = 1-sep-2026: antes solo hay registros sueltos (19/may,
19/ago) que bajaban el promedio (p. ej. SS a ~2,9/día) — decisión de
Mariana.

### Disponibilidad de técnicos (pestaña "Técnicos")

Las horas disponibles de cada técnico se calculan con un horario fijo por
día de la semana (confirmado por Mariana), definido en `HORAS_POR_DIA`
dentro de `renderTecnicos()`:
- Lunes a jueves: 8 h/día
- Viernes: 7 h/día
- Sábado: 3 h/día
- Domingo: 0 h/día (no laborable)

Suma 42 h/semana. `horasDisponiblesRango(desde, hasta)` suma las horas
hábiles de cada día calendario dentro del rango efectivo seleccionado
(recorre día por día, no promedia) — así un rango de un solo día da
exactamente la hora hábil de ese día de la semana (8, 7, 3 o 0 h), y un
rango de varias semanas da el total correcto sin importar en qué día de
la semana empieza o termina.

**Horas extra (domingo):** eliminado (29-sep-2026). Las horas de domingo
suman al tiempo trabajado como cualquier otro día; el domingo sigue sin
sumar al tiempo programado (`HORAS_POR_DIA[0] = 0`).

**Tiempo trabajado SIN tope** (decisión de Mariana, 29-sep-2026, reemplaza
al tope anterior por jornada): cada OT suma su duración real completa (Fecha
Fin Real − Fecha Inicio Real) al técnico, aunque pase de la jornada o de 42
h/semana — "si trabaja más horas también puede salir". La ocupación puede
pasar de 100%. Se le mostró el efecto antes de decidir: con tope 676 h, sin
tope 913 h (29-sep-2026), porque las OT abiertas varios días suman también
las noches (OT 000130 = 169 h de Juan José Mosquera). Las OT de domingo
también suman al tiempo trabajado: ya **no existe** la columna "Horas extra
(domingo)" (Mariana: "todo debe ir en tiempo trabajado"). "Horas libres" no baja de 0. El "Tiempo programado"
sigue siendo el horario legal (`HORAS_POR_DIA`, 42 h/semana).

**Rango en "Todo el periodo":** sin "Desde" elegido, `rangoEfectivoOT()`
arranca en `INICIO_USO_CMMS` (1-sep-2026) y `renderTecnicos()` solo cuenta
OT creadas desde esa fecha — horas disponibles y trabajadas sobre el mismo
periodo (antes una OT suelta del 19/may inflaba las horas disponibles).

**Promedios por técnico** (cuadros arriba de la pestaña): "OT realizadas
por técnico" = promedio de la columna "OT con duración"; "Ocupación por día
por técnico" = horas trabajadas (domingos incluidos) / días laborables (L–S) del
rango, en h/día. Ambos solo sobre técnicos internos que aparecen en alguna
OT del histórico (los de TECNICOS.xlsx sin ninguna OT no bajan el promedio).

**"Horas laborales vs. horas trabajadas por semana"** (pedido de Mariana,
1-oct-2026; pestaña Técnicos, debajo de "Capacidad y ocupación"). En la
pestaña es una **gráfica** (`graficaHorasSemanales()`, Mariana pidió cambiar
la tabla por gráfica): por técnico y semana, barra de horas trabajadas con
línea negra en las laborales (escala hasta 150%; si pasa, barra llena con
borde y el número real). La versión en tabla (`tablaHorasSemanales()`)
queda solo en el modal "Ver detalle por semana" del Resumen. Ambas usan
`datosHorasSemanales()`: filas = técnicos internos (mismas `capRows`),
columnas = semanas lunes–domingo del rango, cada celda = horas trabajadas y %
sobre las laborales de esa semana (42 h, o `horasDisponiblesRango` si la
semana queda cortada por el rango), más Laborales / Trabajadas / Diferencia y
fila de total. Cada OT suma completa en la semana de su Fecha Inicio Real
(fuera del rango → primera/última semana), así los totales coinciden con
"Capacidad y ocupación". Colores del %: >110% rojo, <50% ámbar.
También en el **Resumen, fila "Trabajo"** (junto a la tarjeta Técnicos,
`chartHorasTecnicos()`): gráfico por técnico con barra de horas trabajadas y
línea vertical en sus horas laborales del rango (mismos `tecConsiderados` y
números de la tarjeta Técnicos), y botón "Ver detalle por semana" que abre
`tablaHorasSemanales()` en un modal.

**No volver a un modelo de horas uniformes por día** (p. ej. dividir 42h
entre 6 o 7 días parejo) — el horario real no es uniforme (8h L-J, 7h V,
3h S), y el domingo no suma al tiempo programado.

### Técnicos internos vs. proveedores vs. otros (pestañas "Técnicos" y "Órdenes de Trabajo")

**`TECNICOS.xlsx`** es el maestro de técnicos de planta (columnas NOMBRE,
CEDULA, CIUDAD, hoja única) y **`Datos Generales de Proveedores.xlsx`** es
el maestro de proveedores del CMMS (hoja única `Sheet`, columna `Nombre`).
Ambos viven en la carpeta raíz del proyecto (o en `datos_nuevos/` si
Mariana los pone ahí) y son **opcionales**: si `process_dashboard.py` no
los encuentra, avisa en consola y degrada sin fallar (ver más abajo).

Cada nombre que aparece en el campo `Ejecutores` de una OT se clasifica en
uno de tres grupos, calculados en el frontend (`index_template.html`,
funciones `esOTProveedor()`/`esOTOtro()` en la pestaña "Órdenes de Trabajo"
y las mismas reglas replicadas en `renderTecnicos()` vía `proveedoresSetOf()`
+ `esProveedorRegistrado`):

1. **Técnico interno**: está en `TECNICOS.xlsx` (o en `ALIAS_TECNICOS`, ver
   abajo). Si `TECNICOS.xlsx` no se encuentra, se trata a todos como
   internos (no se separa a nadie).
2. **Proveedor**: no es interno, pero su nombre aparece **tal cual** en
   `Datos Generales de Proveedores.xlsx`. Si ese archivo no se encuentra,
   este grupo queda vacío — nadie se reconoce como proveedor registrado.
3. **Otro**: no es interno y tampoco aparece en el maestro de proveedores
   ni en `PROVEEDORES_EXTRA`. Hoy no hay nadie en este grupo.

**`PROVEEDORES_EXTRA`** (en `process_dashboard.py`, junto a
`cargar_proveedores`): personas que ejecutan OT a nombre de un proveedor sin
estar dadas de alta individualmente en el maestro. Decisión de Mariana
(29-sep-2026): `SEBASTIAN BUITRAGO GRACIANO` cuenta como **proveedor** (antes
estaba en "Otros"). El script las suma a `DATA.proveedores`. Si el aviso de
consola muestra otro nombre "no aparece en Datos Generales de Proveedores ni
en PROVEEDORES_EXTRA", preguntarle a Mariana si es proveedor y agregarlo ahí
(no a `TECNICOS.xlsx` ni a `ALIAS_TECNICOS`).

**Técnicos inactivos:** `TECNICOS.xlsx` trae una columna `ESTADO`
(ACTIVO/INACTIVO). INACTIVO = ya no trabaja en la compañía
(`cargar_tecnicos_inactivos()` → `DATA.tecnicos_inactivos`). Siguen siendo
técnicos internos: su trabajo pasado se muestra en "Capacidad y ocupación"
(con etiqueta "Inactivo"), pero no aparecen en "Ubicación del técnico" ni en
el conteo "Técnicos en sitio / programados" del Resumen
(`tecnicosInternosList()` los excluye).

**No usar "no está en TECNICOS.xlsx" como sinónimo de "proveedor"** — esa
fue la definición vieja de "tercero" (una sola tabla); ahora hay que
distinguir proveedor registrado vs. "otro" usando también el maestro de
proveedores.

En la pestaña **"Técnicos"**, los grupos 2 y 3 se excluyen de "Capacidad y
ocupación por técnico" y de "Dónde está cada técnico ahora" (no tiene
sentido medirles % de ocupación contra 42h/semana ni preguntarse "dónde
está" — no son personal de planta), y en su lugar aparecen en una sola
tabla, "Servicios de proveedores" (OT totales trabajadas, horas trabajadas,
filtro de OT trabajada, duración de esa OT).

En la pestaña **"Órdenes de Trabajo"**, debajo de "Listado completo de
Órdenes de Trabajo" está "Listado completo de Órdenes de Trabajo con
Proveedores" (mismas columnas; OT con algún ejecutor del grupo 2 o 3).

**No hay tablas "Otros"** (Mariana pidió quitarlas el 29-sep-2026, en
Técnicos y en Órdenes de Trabajo): si algún ejecutor cae en el grupo 3, se
muestra en las tablas de proveedores (y en la tarjeta "OT de proveedores" del
Resumen, en la columna Proveedor del listado) con la etiqueta "No
registrado", para no perderlo de vista.

**`ALIAS_TECNICOS`** (en `process_dashboard.py`, junto a `cargar_tecnicos`)
existe porque el nombre del técnico en las OT (`Ejecutores`) a veces trae
más apellidos/nombres que el registrado en `TECNICOS.xlsx` — p. ej.
`TECNICOS.xlsx` trae "ALEJANDRO CARDONA CARMONA" pero en las OT aparece
"ALEJANDRO DE JESUS CARDONA CARMONA". Sin este alias, ese técnico interno
quedaría mal clasificado como tercero. El script imprime en consola
`Nombres en OT no reconocidos como internos` en cada corrida — **revisar
esa lista después de actualizar los datos**: si aparece ahí un nombre que
en realidad es un técnico de planta con el nombre escrito distinto (y no
un tercero real), hay que agregarlo a `ALIAS_TECNICOS`, no ignorarlo.

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
├── TECNICOS.xlsx              <- maestro de técnicos de planta (opcional, cambia poco)
└── datos_nuevos/              <- Mariana coloca aquí los 4 Excel de cada actualización
```

**No editar `index.html` ni `data.json` a mano** — siempre regenerarlos
corriendo `process_dashboard.py`. Si hay que cambiar el diseño, los
colores, o agregar una pestaña, el cambio va en `index_template.html`
(HTML/CSS/JS) y luego se vuelve a inyectar el JSON.

## Identidad de marca

Colores extraídos del tablero de referencia de la empresa (paleta ya
aplicada en `index_template.html`, no recalcular). Paleta corporativa
(pedido de Mariana, 1-oct-2026: "muchos colores, que se vea más
profesional"): vino de la marca + neutros cálidos; verde/ámbar/rojo
apagados **solo** para semáforos (alertas, cumplimiento, vencidas). Todas
las tarjetas destacadas (`accent-c1/c2/c3` y `good`) llevan borde y número
en vino. No volver a azul/naranja/verde/morado en tarjetas ni gráficas.
```
--maroon-dark:#52101E;  --maroon:#7A1A2E;  --cream:#F8F3EE;
--good:#3F7D5C; --warn:#B7832F; --bad:#B3261E;
--chart-1:#7A1A2E (vino, Correctivo)  --chart-2:#A08463 (bronce, Preventivo)
--chart-3:#5B6770 (pizarra)  --chart-4:#B9737E  --chart-5:#3E4A52  --chart-6:#9A8F8B
```
Los colores por tipo de las gráficas salen de `tiposYColores()`: Correctivo
y Preventivo fijos, los demás tipos toman el siguiente `--chart-N` libre.
