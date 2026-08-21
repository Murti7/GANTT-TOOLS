# Project Structure

Canonical future structure:

```text
projects/<project-slug>/
  project.yaml
  input/
    budget/*.bc3
    planning/planning.yaml
    billing/
  output/
    budgeting/
    planning/
    billing/
```

Legacy structure remains supported:

```text
projects/<project-slug>/input/
  *.bc3
  config.yaml
  planificacion.yaml
```

When `project.yaml` is absent, `config.yaml` is adapted to `ProjectModel` and a
traceable warning is added to execution context/manifest.
