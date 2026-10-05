# Excel Template Export (Obsidian)

Plugin de Obsidian que llena templates de Excel (`.xlsx`) con las properties (frontmatter) de tus notas, **preservando el formato del template**: colores, anchos, celdas combinadas, logos y fórmulas.

Sirve para:

- **Una nota puntual:** exportar la nota activa.
- **Un conjunto filtrado:** todas las notas que cumplan un filtro.
- **Relaciones entre notas (joins):** traer las notas que apuntan a la raíz (backlinks) y navegar links hacia adelante. Por ejemplo: Proceso → Interviews → Persona.

> Estado: v0.1 (desktop). El core no depende de Node/Electron, así que mobile se puede habilitar más adelante.

---

## Instalación (manual)

1. `npm install && npm run build`
2. Copiá `main.js`, `manifest.json` y `styles.css` a `<vault>/.obsidian/plugins/excel-template-export/`.
3. Activá el plugin en *Settings → Community plugins*.

En [`examples/`](examples/) hay dos templates (`Proceso.xlsx`, `Consolidado.xlsx`) y sus definiciones listas para copiar al vault.

## Cómo funciona

1. Diseñás un Excel normal y ponés **placeholders** en las celdas: `{{proceso.role}}`.
2. Creás una **definición** (una nota con un bloque `excel-export`) que dice cuáles son las notas raíz, cómo se filtran, qué notas relacionadas traer y qué template usar.
3. Ejecutás el export desde el bloque, la paleta de comandos o el menú contextual de una nota.

### Definición

Comando **Excel Export: Nueva definición** crea una nota de ejemplo en la carpeta de definiciones (default `Exports/`).

````markdown
```excel-export
name: Procesos activos                  # requerido
template: Templates/Excel/Proceso.xlsx  # requerido, path dentro del vault

root:
  alias: proceso                        # nombre usado en los placeholders
  where: 'categories contains "Talent Acquisition Process"'   # universo de candidatas
  filter: 'status != "Closed"'          # opcional; se ignora en "Exportar nota activa"
  sort: 'created desc'                  # opcional: campo asc|desc, separados por coma

mode: file-per-root                     # file-per-root | sheet-per-root | single

relations:
  interviews:                           # backlink: Interview.process apunta a la raíz
    from: 'categories contains "Interview Evaluations"'
    on: 'process -> proceso'
    sort: 'date asc'
  hires:                                # relación derivada de otra
    source: interviews
    filter: 'decision == "Hire"'

output:
  folder: Exports/out
  filename: '{{proceso.file.name}} - {{@today | date:"yyyy-MM-dd"}}.xlsx'
  sheetName: '{{proceso.file.name}}'    # solo sheet-per-root
  overwrite: ask                        # ask | overwrite | suffix

normalize:                              # opcional; si falta se usan los settings
  links: display                        # display | target | raw
  listSeparator: ', '
  stars: number                         # number | raw
  emptyValue: ''

emptyBlock: remove                      # opcional: remove | blank (fila #each sin elementos)
```
````

- `where` define el universo y `filter` lo restringe. "Exportar nota activa" valida contra `where` e ignora `filter`.
- `on: <campo> -> <alias>`: el campo (link o lista de links) de la candidata apunta a la raíz. Se compara por **archivo resuelto**, no por texto.
- Los links hacia adelante **no se declaran**: se navegan en el path (`interviews.interviewed.linkedin`).
- Un YAML inválido muestra el error en el propio bloque, con la clave o la columna del problema.

### Template de Excel

| Sintaxis | Qué hace |
|---|---|
| `{{proceso.role}}` | Escalar. Si la celda tiene **solo** el placeholder, se escribe el valor tipado (número, fecha, booleano) y se respeta el formato de número de la celda. Mezclado con texto (`Rol: {{proceso.role}}`) da texto. |
| `{{interviews.tech rating}}` | Los nombres de property pueden tener espacios. |
| `{{proceso["mi.prop"]}}` | Corchetes para nombres con puntos. |
| `{{interviews.interviewed.seniority}}` | Si un segmento es un link, se resuelve y se navega. Si es una lista de links, se mapea. |
| `{{proceso.file.name}}` | Pseudo-properties: `file.name`, `file.path`, `file.folder`, `file.ext`, `file.ctime`, `file.mtime`, `file.link`. |
| `{{@index}}`, `{{@count}}`, `{{@today}}`, `{{@now}}`, `{{@exportName}}` | Variables especiales (`@index` es 1-based dentro de un `#each`). |
| `{{#each interviews}}…` | Al **inicio** de una celda: marca la fila como plantilla y la repite una vez por elemento (copiando estilos, alto y merges horizontales). Un `#each` por fila; puede haber varias filas `#each` por hoja. |

Dentro de una fila `#each interviews`, `{{interviews.decision}}` es el elemento actual y los campos de la raíz siguen disponibles (`{{proceso.role}}`).

#### Pipes

`{{path | pipe:arg | pipe}}`

| Pipe | Efecto |
|---|---|
| `raw` | Sin normalizar (mantiene `[[...]]`). |
| `target` | Nombre de la nota destino en vez del alias. |
| `join:"; "` | Separador de lista. |
| `first` | Primer elemento de una lista. |
| `count` | Tamaño de una lista o colección: `{{hires \| count}}`. |
| `stars` | Cuenta ⭐ (½ suma 0.5) y devuelve un número. |
| `date:"dd/MM/yyyy"` | Formatea como texto (tokens `yyyy yy MM M dd d HH H mm ss`). Sin pipe, las fechas van como fecha de Excel. |
| `default:"-"` | Valor si está vacío. |
| `upper` / `lower` | Mayúsculas / minúsculas. |
| `link` | Hipervínculo a `obsidian://open?...` con el texto normalizado. |

### Modos

| Modo | Resultado |
|---|---|
| `file-per-root` | Un `.xlsx` por nota raíz. |
| `sheet-per-root` | Un `.xlsx`; la primera hoja del template se clona por raíz (nombre de hoja sanitizado a 31 caracteres y sin repetir). Las demás hojas se mantienen y ven la colección de raíces (`{{proceso \| count}}`, `{{#each proceso}}`). |
| `single` | Un `.xlsx`; la colección de raíces se llama como el alias: `{{#each proceso}}`. **Flatten:** `{{#each proceso.interviews}}` da una fila por interview, con acceso al padre (`{{proceso.role}}`) y al elemento (`{{interviews.interviewed}}`). Dentro de `{{#each proceso}}` también están las relaciones de cada raíz (`{{interviews \| count}}`). |

### Filtros

```
expr  := or ;  or := and ("or" and)* ;  and := unary ("and" unary)*
unary := "not" unary | "(" expr ")" | path op value | path ("exists" | "!exists")
op    := == | != | > | >= | < | <= | contains | !contains
value := "texto" | 123 | true | false | null | today | date("2026-03-01") | alias
```

- **Links:** se comparan contra el nombre de la nota destino, sin distinguir mayúsculas. `role == "Backend"` matchea `[[Backend]]` y `[[Roles/Backend|Back]]`.
- **Listas:** `contains` verifica pertenencia; `==` es verdadero si algún elemento coincide.
- **Texto:** `contains` es substring sin distinguir mayúsculas.
- **Fechas:** ISO (`YYYY-MM-DD`, hora opcional); `today` es la fecha local.
- **Campos faltantes:** toda comparación es falsa, salvo `!=`, `!contains` y `!exists`. `campo == null` es verdadero si falta.
- Properties con espacios: `["tech rating"] >= 3`. Los paths pueden navegar links: `interviewed.seniority == "Junior"`.
- En los filtros de una relación, el alias de la raíz se puede usar como valor: `process == proceso`.

### Normalización de valores

| Entrada | Salida (default) |
|---|---|
| `[[Backend]]` | `Backend` |
| `[[Celula X\|Alias]]` | `Alias` (`links: target` → `Celula X`) |
| `[[Carpeta/Nota]]` | `Nota` |
| lista | elementos normalizados unidos con `listSeparator` |
| `2026-03-01` | fecha de Excel (formato de la celda o `yyyy-mm-dd`) |
| `⭐⭐⭐` | `3` |
| número / booleano | tipo nativo |
| vacío / null | `emptyValue` |
| objeto | JSON compacto + warning |

Las fechas se interpretan en hora local, sin corrimiento por zona horaria.

## Comandos

- **Ejecutar export…**: elegís una definición y la corre.
- **Exportar nota activa con…**: muestra las definiciones cuyo `root.where` coincide con la nota activa (si ninguna coincide, muestra todas con aviso) y exporta solo esa nota. También está en el menú contextual de las notas: *Exportar a Excel con…*.
- **Validar template…**: lista placeholders y filas `#each`, colecciones desconocidas y paths que no resuelven contra la primera raíz.
- **Nueva definición**.

El bloque `excel-export` muestra nombre, template, modo, cantidad de raíces y los botones **Ejecutar** y **Validar**.

Después de exportar aparece un aviso (`3 archivos generados, 2 warnings`) con **Ver detalle** (placeholders sin resolver, links rotos, notas omitidas, fórmulas) y **Abrir carpeta**. Los placeholders sin resolver nunca abortan el export.

## Settings

| Setting | Default |
|---|---|
| Carpeta de definiciones | `Exports/` |
| Carpeta de salida por defecto | `Exports/out/` |
| Links | `display` |
| Separador de listas | `, ` |
| Estrellas | `number` |
| Valor vacío | `""` |
| Fila `#each` vacía | `remove` |
| Abrir archivo al terminar (desktop) | `false` |
| Sobrescritura por defecto | `ask` |

## Limitaciones conocidas (v1)

- **Fórmulas:** al insertar filas se ajustan las referencias A1 de la misma hoja (y las de otras hojas con prefijo `Hoja!`): las de abajo se desplazan y los rangos que contienen la fila plantilla se expanden (`AVERAGE(G7:G7)` → `AVERAGE(G7:G9)`). Si la fila `#each` queda sin elementos y se elimina, las referencias a esa fila quedan en `#REF!` (como en Excel) y se avisa en el reporte; usá `emptyBlock: blank` para conservarla. Referencias de filas completas (`7:7`), nombres estructurados y fórmulas matriciales no se ajustan.
- Conviene ubicar el bloque `#each` al final de la hoja.
- **Formato condicional y validaciones de datos** no se desplazan (se avisa).
- **Celdas combinadas** que cruzan verticalmente una fila `#each` se descombinan (con warning).
- **Tablas de Excel (ListObjects):** no se redimensionan; en `sheet-per-root` no se copian a las hojas clonadas (warning).
- Un solo nivel de `#each` por fila; bloques de varias filas no están soportados (el caso consolidado se cubre con flatten).
- `on` solo admite `<campo> -> <aliasRaíz>`.
- No se evalúan archivos `.base`; los filtros usan su propio lenguaje.

## Decisiones sobre las preguntas abiertas (§13 del spec)

1. La property de categoría puede ser texto o wikilinks: los filtros comparan links por nombre de nota destino, así que `categories contains "TH"` funciona en ambos casos.
2. "Activo" = `status != "Closed"` (un `status` faltante cuenta como activo).
3. Estrellas: se cuenta ⭐ (con o sin selector de variación), `½` suma 0.5 y los números pasan tal cual.
4. Tablas de Excel: ver limitaciones.
5. Mobile: el core es puro, pero el manifest declara `isDesktopOnly: true` hasta probarlo.

## Desarrollo

```bash
npm install
npm test          # Vitest, en Node, sin Obsidian
npm run build     # typecheck + bundle (main.js)
npm run dev       # watch
```

```
src/
  main.ts, settings.ts      # plugin, comandos, settings (Obsidian)
  adapter/                  # VaultAdapter: Obsidian + memoria (tests)
  model/                    # NoteRecord, reporte de warnings
  config/                   # schema, validación y parseo del YAML
  query/                    # lexer, parser, evaluador y sort de filtros
  graph/                    # paths, relaciones (índice inverso), runtime
  values/                   # links, normalización, pipes
  excel/                    # placeholders, scanner, motor, fórmulas, IO
  export/                   # runner (modos y escritura) y validación
  ui/                       # picker, code block, modales
tests/                      # tests del core + fixture del dominio (vault.json)
```

El core (`query`, `graph`, `values`, `excel`, `export`, `config`) no importa `obsidian`; Obsidian entra solo por `adapter/obsidian-adapter.ts` y `ui/`.
