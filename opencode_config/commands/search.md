# /search

Web search via Brave API using Python requests. **Fallback when TUN is OFF**.

## When to use

| State | Preferred tool | Why |
|-------|---------------|-----|
| TUN ON | Built-in `brave-search` MCP | Native, fast, structured |
| TUN OFF | `/search` command | Node.js can't reach Brave API; Python reads system proxy |

## Usage

```
/search <query> [--count N]
```

## Implementation

```bash
python "%USERPROFILE%\.config\opencode\scripts\web_tools.py" search <query> [--count <n>]
```
