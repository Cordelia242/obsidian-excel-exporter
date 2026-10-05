Una fila por interview de cada proceso activo. Abrí **Setup** para ver o cambiar qué columna se llena con qué dato.

```excel-export
name: Postulantes en procesos activos
template: Templates/Excel/Consolidado.xlsx
root:
  alias: proceso
  where: categories contains "Talent Acquisition Process"
  filter: status != "Closed"
  sort: priority desc, created asc
mode: single
relations:
  interviews:
    from: categories contains "Interview Evaluations"
    on: process -> proceso
output:
  folder: Exports/out
  filename: 'Consolidado {{@today | date:"yyyy-MM-dd"}}.xlsx'
cells:
  Consolidado!A2: proceso.file.name
  Consolidado!B2: proceso.role
  Consolidado!C2: proceso.priority
  Consolidado!D2: interviews.interviewed
  Consolidado!E2: interviews.tech rating | stars
  Consolidado!F2: interviews.decision
rows:
  Consolidado!2: proceso.interviews
```
