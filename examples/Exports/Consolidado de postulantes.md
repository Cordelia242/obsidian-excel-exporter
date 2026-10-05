Una fila por interview de cada proceso activo (§14.2).

```excel-export
name: Postulantes en procesos activos
template: Templates/Excel/Consolidado.xlsx
root:
  alias: proceso
  where: 'categories contains "Talent Acquisition Process"'
  filter: 'status != "Closed"'
  sort: 'priority desc, created asc'
mode: single
relations:
  interviews:
    from: 'categories contains "Interview Evaluations"'
    on: 'process -> proceso'
output:
  folder: Exports/out
  filename: 'Consolidado {{@today | date:"yyyy-MM-dd"}}.xlsx'
```
