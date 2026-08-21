# Planning

Input:

- BC3 file.
- `planning.yaml`.

Output in ENG-3 capability API:

```text
output/planning/<source-id>/
  planning_analysis.xlsx
  gantt.png
  run_manifest.json
```

Planning calculates durations, scenarios, resource load, critical path and
diagram data from domain models. It does not depend on budget Excel outputs.
