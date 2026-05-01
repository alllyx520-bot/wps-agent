---
description: 分析项目结构，生成完美的 .gitignore，优化 opencode 在大项目中的性能
agent: build
---

Analyze the current project root and generate or improve `.gitignore` to:

1. Identify project type by reading key files (`package.json`, `requirements.txt`, `pyproject.toml`, `Cargo.toml`, etc.)
2. Scan the directory tree for files/directories that should be ignored
3. Read existing `.gitignore` if present
4. Merge and deduplicate with existing entries
5. Sort logically
6. Write the final `.gitignore` file
7. Verify the file was written correctly
8. Summarize what was added and why

Use glob and read tools to inspect the project. Do not guess — actually scan the directory. Always preserve existing entries that are still relevant.
