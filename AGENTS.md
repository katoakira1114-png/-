# AGENTS.md - Instructions for Autonomous & Pair AI Agents

Please refer to and strictly follow [AI_COLLABORATION_GUIDELINES.md](./AI_COLLABORATION_GUIDELINES.md) and [.cursorrules](./.cursorrules).

## Critical Guidelines
1. **Never rollback or delete existing functionality**: Multiple developers and AI assistants collaborate on this repository. Treat all unfamiliar code as deliberate contributions. Implement new requests additively.
2. **Audit trail & timestamps**: When editing files, update the `@lastModified`, `@modifiedBy`, and `@note` JSDoc/HTML header comments, and update `SYSTEM_META` in `無題.js`.
3. **Schema & OAuth safety**: Do not alter sheet names or column structures in Google Spreadsheet `1ZiNel8TGJ-e7JMolkRagVOSnr1UczvQzpTa-F-s4p3o`. Do not alter `appsscript.json` OAuth scopes or access configurations.
4. **No hardcoded secrets**: Secrets must be read from Script Properties (`PropertiesService.getScriptProperties()`) or environment variables.
