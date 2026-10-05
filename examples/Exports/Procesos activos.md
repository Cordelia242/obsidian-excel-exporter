Ficha por proceso con sus postulantes. El Excel (`Templates/Excel/Proceso.xlsx`) solo tiene títulos y formato: qué dato va en cada celda se configura con el botón **Setup**.

```excel-export
name: Procesos activos
template: Templates/Excel/Proceso.xlsx
root:
  alias: proceso
  where: categories contains "Talent Acquisition Process"
  filter: status != "Closed"
  sort: created desc
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
  filename: '{{proceso.file.name}} - {{@today | date:"yyyy-MM-dd"}}.xlsx'
  overwrite: ask
cells:
  Proceso!B1: proceso.file.name
  Proceso!B2: proceso.team
  Proceso!E2: proceso.role
  Proceso!B3: proceso.seniority
  Proceso!E3: proceso.quantity
  Proceso!B4: proceso.status
  Proceso!E4: proceso.hires | count
  Proceso!A7: interviews.interviewed
  Proceso!B7: interviews.interviewed.seniority
  Proceso!C7: interviews.interviewed.stack
  Proceso!D7: interviews.interviewed.linkedin
  Proceso!E7: interviews.date
  Proceso!F7: interviews.speech rating | stars
  Proceso!G7: interviews.tech rating | stars
  Proceso!H7: interviews.decision
rows:
  Proceso!7: interviews
```
