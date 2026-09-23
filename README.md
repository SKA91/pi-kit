# pi-kit

Personal pi coding agent extensions and sbx mixin kit (SKA91).

## Layout

```
package.json               root pi manifest — loads extensions/ when the repo
                           is installed as a git package (pi install git:...)
extensions/                my pi extensions (one package dir per extension)
  zeldoc-web-search/       vendored Zeldoc.ai web search (web_search +
                           fetch_content), from docs.zeldoc.ai/web-search
  calendar-widget/         O365 calendar widget below the editor (next two
                           meetings; Graph API via O365_* env vars, "~"
                           until the first fetch succeeds)
kit/eet-personal/          sbx mixin kit installing everything into sandbox
                           VMs (see its README)
```

## Quick start (sandbox)

```bash
sbx kit add eet ~/projects/github.com/ska91/pi-kit/kit/eet-personal/
```

See `kit/eet-personal/README.md` for details, create-time usage, and the
extension development loop.
