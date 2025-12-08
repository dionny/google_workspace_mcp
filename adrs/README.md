# Architecture Decision Records (ADRs)

This directory contains Architecture Decision Records (ADRs) for the Google Workspace MCP project.

## What is an ADR?

An Architecture Decision Record (ADR) is a document that captures an important architectural decision made along with its context and consequences.

## ADR Format

Each ADR should follow this template:

```markdown
# [Number]. [Title]

Date: YYYY-MM-DD

## Status

[Proposed | Accepted | Deprecated | Superseded by ADR-XXX]

## Context

What is the issue that we're seeing that is motivating this decision or change?

## Decision

What is the change that we're proposing and/or doing?

## Consequences

What becomes easier or more difficult to do because of this change?

### Positive Consequences

- [benefit 1]
- [benefit 2]

### Negative Consequences

- [drawback 1]
- [drawback 2]

## Alternatives Considered

What other options were considered and why were they not chosen?

## Related

- Links to related ADRs, tickets, or documentation
```

## Naming Convention

ADRs are numbered sequentially and stored as markdown files:

```
adrs/
  001-use-fastmcp-framework.md
  002-auto-clean-blank-lines-for-lists.md
  003-...
```

## Index

| Number | Title | Status | Date |
|--------|-------|--------|------|
| 001 | [Auto-clean blank lines when converting to lists](001-auto-clean-blank-lines-for-lists.md) | Accepted | 2024-12-08 |

## Resources

- [ADR GitHub Organization](https://adr.github.io/)
- [ADR Tools](https://github.com/npryce/adr-tools)

