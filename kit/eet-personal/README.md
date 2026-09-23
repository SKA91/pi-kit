# kit/eet-personal

Personal sbx mixin kit: installs my pi extensions into the sandbox VM's
global pi scope (`~/.pi/agent`), so they survive `sbx stop/start` and are
re-applied automatically on sandbox create/recreate. Personal — colleagues
have their own VMs and never see this.

## What it installs (pi packages, via `pi install`)

| Package | Source | Purpose |
|---|---|---|
| `npm:@juicesharp/rpiv-ask-user-question@2.11.0` | community | structured questionnaire tool |
| `npm:@juicesharp/rpiv-todo@2.11.0` | community | live todo overlay |
| `git:github.com/SKA91/pi-kit@main` | mine | extensions in `extensions/` |

Plus the vendored Zeldoc.ai web-search extension (see below).

`pi-web-access` was tried and removed: it registers `web_search` and
`fetch_content` too, and pi refuses to load both it and the Zeldoc
extension (one search extension max). The Zeldoc one won — it needs no
separate search-provider account, only the Zeldoc key pi already uses.

## Zeldoc.ai web search (vendored)

`extensions/zeldoc-web-search/zeldoc_search.ts` is vendored from
<https://docs.zeldoc.ai/search/pi/zeldoc_search.ts>. It adds `web_search`
(private SearXNG search via `api.zeldoc.ai`, authenticated with the same
Zeldoc key pi already uses through the sandbox proxy) and `fetch_content`
(local page fetch as readable text).

## Applying

Existing sandbox:

```bash
sbx kit add eet ~/projects/github.com/ska91/pi-kit/kit/eet-personal/
```

At create/recreate (together with the team kit):

```bash
DOCKER_SANDBOXES_DOCKER_SIZE=40g sbx create --name eet shell . \
  --kit sbx-kit \
  --kit ~/projects/github.com/ska91/pi-kit/kit/eet-personal/
```

Any other sandbox works the same way — pass the kit dir to `sbx create
--kit` or `sbx kit add <sandbox>`.

## Updating my extensions

1. Edit/add under `extensions/`, commit, push to `main`.
2. In the VM: `pi update git:github.com/SKA91/pi-kit` (reconciles the clone
   to the ref and reloads), or bump the ref in `spec.yaml` and re-apply the
   kit for recreate-time determinism.

## Developing an extension (VM-local loop)

The host repo is not mounted in the VM, so for live development clone
inside the VM and install as a local-path package (pi does not copy local
paths, so edits are picked up by `/reload`):

```bash
ssh eet.sbx
git clone https://github.com/SKA91/pi-kit ~/pi-kit
pi install ~/pi-kit/extensions/<my-ext>
# edit ~/pi-kit/extensions/<my-ext>/... (or edit on host + git pull in VM), then /reload in pi
```

Promote to everyone-sandbox-installed by pushing and re-applying the kit.

## Note on `main` pinning

The kit pins `@main` for now (keyless public clone, always-current).
Trade-off: recreates re-clone whatever `main` is at that moment. For
full determinism, push a tag and pin `@<tag>` instead — pi skips pinned
git refs on `pi update`, so moving requires bumping the ref explicitly.
