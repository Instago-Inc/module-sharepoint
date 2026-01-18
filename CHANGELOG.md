## v1.0.1 - Modern SharePoint helper docs
Fresh documentation now spotlights SharePoint, OneDrive, and Microsoft Graph automation patterns with workflow-friendly examples. Updating the dependency calls to latest versions keeps the helper aligned with shared Graph utilities and reduces pinning drift.

### Added
- A technical reference that explains exports, usage scenarios, and configuration for the SharePoint helper.

### Changed
- Require the latest versions of `b64`, `json`, `path`, `fmt`, `graph`, and `log` so hosted workflows can pick up fixes from those packages.
