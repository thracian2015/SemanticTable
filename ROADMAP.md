# Roadmap

Semantic Table is currently a public beta. Priorities may change as Excel and semantic-model compatibility is validated.

## Beta stabilization

- Validate connection discovery and authenticated metadata access across supported Microsoft 365 builds.
- Add automated tests for DAX generation, state migration, and filter serialization.
- Improve diagnostics and actionable connection errors.
- Verify behavior under Power BI row-level security.

## Field and query experience

- Add drag-and-drop ordering for selected fields.
- Add structured server-side sorting and additional filter types.
- Define model-specific grain keys for detail-row extracts.
- Improve calculation-group handling.

## Distribution readiness

- Establish repeatable signed Release builds.
- Document trusted-location and endpoint-management deployment options.
- Add release packaging and upgrade guidance.
