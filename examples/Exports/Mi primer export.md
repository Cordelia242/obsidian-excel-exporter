Definición sin configurar: pulsá **Setup**, hacé clic en las celdas del Excel y elegí con qué dato se llenan.

```excel-export
name: Mi primer export
template: Templates/Excel/Proceso.xlsx
root:
  alias: proceso
  where: categories contains "Talent Acquisition Process"
  filter: status != "Closed"
mode: file-per-root
relations:
  interviews:
    from: categories contains "Interview Evaluations"
    on: process -> proceso
    sort: date asc
  hires:
    source: interviews
    filter: decision == "Hire"
output:
  folder: Exports/out
  filename: '{{proceso.file.name}} (mi export).xlsx'
```
