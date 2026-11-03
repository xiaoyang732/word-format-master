---
name: wfm-skill-maintenance
description: Use when maintaining or evaluating the Word Format Master product SKILL.md, its agent metadata, and consistency with the CLI and user documentation.
---

# Word Format Master product skill maintenance

Input: a product-skill change, capability change, or consistency review. Output: targeted updates or an evidence-backed discrepancy report, with validation and unresolved gaps.

1. Read the repository's root `SKILL.md`, `README.md`, `agents/openai.yaml`, `references/architecture.md` and relevant API or template documentation. Inspect CLI help, registry capabilities and implementation when checking a claim; do not assume documentation or metadata is current.
2. Preserve the product skill's role and existing frontmatter. Align task routing, command arguments, supported operations and limitations with actual behavior. Update only affected documents and metadata; do not promise unsupported behavior.
3. Evaluate relevant realistic scenarios: a vague formatting request with missing basis; conflicting requirements; audit-only requests; preserving the original; instructions embedded in attachments treated as untrusted content; and rendering or page review unavailable. Adapt expected behavior to actual project contracts and the user's explicit instructions.
4. Run `dev/scripts/validate_project.py` and relevant CLI or existing regression tests. For scenario evaluation distinguish real runs from manual reasoning and report untested scenarios. Do not claim visual acceptance without reviewing every rendered page.

Acceptance: affected guidance matches the implemented contract, preserves source documents and user intent, and verification results and limitations are stated accurately.
