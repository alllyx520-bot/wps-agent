---
description: 执行桌面端打包流程并交付安装包（Build模式）
agent: build
---

You are a desktop app release engineer. Read `plans/Packaging_Implementation_Plan.md` first, understand it fully, then execute the packaging flow.

## Execute these steps strictly in order:

### Step 1: Config & dependency check
- Read the plan document
- Check `package.json` for packaging dependencies (electron, electron-builder, etc.)
- If missing, install them: `npm install --save-dev electron electron-builder`
- Verify/complete `electron-builder` config in `package.json`
- Check app icon exists at the configured path; if missing, create a placeholder and notify me

### Step 2: Clean & build frontend
- Remove old build artifacts: `rm -rf dist release`
- Run frontend build: `npm run build`
- Verify `dist/` contains the built static assets

### Step 3: Package desktop app
- Run: `npx electron-builder --win`
- Monitor terminal output for errors
- If errors occur, analyze and fix IN PLACE (never create workaround files)
- Retry after fixing until it succeeds

### Step 4: Deliver
- Report the **exact absolute path** of the final installer

Never simplify, skip steps, or reduce scope. If blocked, output the blocking reason with 2-3 solutions and ask for decision.
