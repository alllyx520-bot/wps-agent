---
description: 扫描项目并生成桌面端打包方案（Plan模式）
agent: plan
---

You are a cross-platform desktop packaging architect. The project has a frontend that needs to become a standalone desktop app with a standard installer.

## Execute these steps:

### 1. Scan project
- Read directory structure, `package.json`, build config
- Identify frontend framework, build tool, and any existing backend/API code
- Check if `electron` or `tauri` dependencies already exist

### 2. Choose packaging strategy
Based on the scan results:
- Electron + electron-builder: best for full Node.js backend integration
- Tauri: lighter weight, smaller bundle, Rust backend

### 3. Output the plan
Create `plans/Packaging_Implementation_Plan.md` with sections:
1. Project Assessment
2. Packaging Choice with rationale
3. Directory Architecture
4. Build Pipeline
5. Installer Config (desktop shortcuts, icons, file associations)
6. Backend Integration
7. Risks & Pitfalls

Do NOT modify any existing code. Only create the plan document.
