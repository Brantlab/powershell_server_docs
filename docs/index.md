# Server Inventory Docs

This site is designed for nightly-published Windows server inventory generated from Active Directory with PowerShell and rendered with Material for MkDocs.

## What You’ll Find

- A generated page for each server with OS, patch visibility, storage, uptime, and network details
- An overview page that can act as a quick operational dashboard
- Material for MkDocs formatting tuned for readability by both technical and non-technical viewers

## Start Here

- Open [Server Overview](servers/index.md) once your first inventory run has generated content
- Preview the layout now with [Sample Server Page](servers/SAMPLE-SRV01.md)

## Publishing Notes

The GitHub Actions workflow publishes this site with `mkdocs gh-deploy --force`. Generated server Markdown should land under `docs/servers/`.
