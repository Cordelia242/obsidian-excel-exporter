Ficha por proceso con sus postulantes (§14.1).

```excel-export
name: Procesos activos
template: Templates/Excel/Proceso.xlsx

root:
  alias: proceso
  where: 'categories contains "Talent Acquisition Process"'
  filter: 'status != "Closed"'
  sort: 'created desc'

mode: file-per-root

relations:
  interviews:
    from: 'categories contains "Interview Evaluations"'
    on: 'process -> proceso'
    sort: 'date asc'
  hires:
    source: interviews
    filter: 'decision == "Hire"'

output:
  folder: Exports/out
  filename: '{{proceso.file.name}} - {{@today | date:"yyyy-MM-dd"}}.xlsx'
  overwrite: ask
```
