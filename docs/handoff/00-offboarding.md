# What breaks when Jason's accounts are disabled

Written 2026-09-10, updated 2026-09-23. About a departure expected around October 2026. Each item says who
has to act, because none of these can be fixed from inside this repo by whoever inherits it.

**Who is taking this over: Atul Joshi** (ClickUp user `113594290`). Pronouns not recorded — use they/them
unless they say otherwise. If you are Atul: read this file, then
[03-running-headless.md](03-running-headless.md), then the runbook for the one open job. The rest can wait
until you need it.

## 1. Repository ownership — transfer to the Task Agency account

The repo was built under Jason's **personal** GitHub account, `jasonweblifestores`, and pushed public. Before
he left he confirmed the destination: the GitHub organisation **`weblife-task-agency`** (display name
"Task Agency", created 2026-08-20). The plan of record is to **transfer this repository into that org** rather
than leave it on a personal account or re-push it somewhere new. After the transfer it becomes
`weblife-task-agency/DocRefinePro`, with a redirect from the old URL.

**Transferring into an org needs repository-creation permission in that org** — write access to one of its
repos is not enough. If the transfer stalls, that is the likely reason, and an org Owner has to grant it.

**If you are reading this and `origin` still points at `jasonweblifestores/DocRefinePro`, the transfer did not
happen — chase it.** Nobody at WebLife can administer the repo while it lives there: no settings, no Actions
secrets, no releases, no CI reruns.

**What a GitHub transfer does and does not carry over** — this matters for the item below:

- **Carried:** full commit history, branches, tags, releases and their assets, issues, PRs, stars and watchers.
  GitHub also leaves a redirect at the old URL, so existing clones and any link in these docs keep working.
- **NOT carried: Actions secrets.** They are scoped to the repository *in its old owner's context* and do not
  survive the move. The `CLICKUP_API_TOKEN` secret will be gone after the transfer and must be re-added by the
  new owner. See item 2.
- Also worth checking after the move: branch protection rules, and that the Actions workflows are enabled at
  all — a transferred repo can land with Actions disabled until an owner turns them on.

**Visibility after the move.** A transfer preserves visibility, so this lands in the org still **public**.
The org's other repositories are private, so if that is the house norm, an owner can flip it after the
transfer — nothing in this pack depends on it being public. Be aware that anything pushed while it was public
is already public regardless.

**Note on the original decision.** Jason's explicit decision (2026-09-10) was to keep the repo public and put the full
handoff in it, so that anyone could clone it later without needing access to anything else. That is why
internal delivery counts, ClickUp task IDs and vendor addresses appear in these docs. It was a deliberate
trade for continuity. The new owner is free to flip it private — everything here works the same either way —
but be aware that anything already pushed while it was public is already public.

## 2. CI's ClickUp integration runs on Jason's personal API token — IT WILL BREAK

`.github/workflows/build.yml` (around line 178) reads `secrets.CLICKUP_API_TOKEN`. That secret is **Jason's
personal ClickUp API token**, added 2026-08-04. When his ClickUp account is deactivated the token stops
working, and the `notify-clickup` job stops creating release subtasks under
[86ex00r23](https://app.clickup.com/t/86ex00r23).

The job is written to fail loudly-ish rather than silently — it writes a warning to the run summary when the
token is not visible — but a revoked-but-present token may behave differently from a missing one. **Test it on
the first release after the handover.**

**Action:** after the repo transfer — which wipes the secret outright, see item 1 — whoever owns releases
generates their own ClickUp personal API token
(ClickUp → avatar → Settings → Apps → API Token) and replaces the repo secret. **No token is stored anywhere
in this repo, by design.** Do not commit one.

## 3. The knowledge in `docs/handoff/knowledge/` was personal, local, and nearly lost

Those eight documents were Claude Code memory files in
`C:\Users\WORK\.claude\projects\...\memory\` — one laptop, no backup, invisible to everyone else. They are in
the repo now. That is the fix, and it is done. Mentioning it because the same thing will happen again to
whoever takes over unless they keep notes somewhere shared.

## 4. `Documents\DocRefinePro_Data\` was never in version control

The parts that could not be regenerated are now in this repo (`brandkits/`, `verify/`,
`docs/handoff/deliveries/`). Deliberately **not** copied, because they are large and reproducible:

- `DocRefinePro_Data\Workspaces\` — 3.0GB of ingest workspaces for the two batches. Regenerable by re-ingesting.
- `DocRefinePro_Data\verify\_p0out\` — 31MB of test output. Regenerated by running the suite.
- Logs (`app_debug.log`, `app_events.jsonl`, `rebrand_runs.jsonl`) — local run history, no lasting value.

If the delivered output trees themselves matter to anyone (`Batch 4\_unique-to-rebrand_rebranded`,
`MBW rebrand work\_unique-to-rebrand_rebranded`), they are **not** here — they are gigabytes. Both were
uploaded to Google Drive at delivery; the MBW deliverables live *inside* the source folder on the shared
drive. Confirm they are still there before Jason's Drive access ends.

## 5. Google Drive — verified intact 2026-09-23, and access arrives with the next batch

Checked on the mapped `G:` drive before handover. Everything is where it should be, under
`G:\Shared drives\[VP] Venia Products KMS\Shared_Services\DP\DP_CONT_Operational\Operational - Task Related Docs & Sheets\`:

| Location | State |
|---|---|
| `BM Downloadable re-branding\Batch 4\_unique-to-rebrand_rebranded` | **2,174 delivered files present** (matches local exactly) |
| `BM Downloadable re-branding\Batch 4\Template` | Kit artwork present — Portrait + `Landscaape` (their spelling), 5 PNGs each |
| `MBW Downloadable re-branding\_unique-to-rebrand_rebranded` | **970 delivered files present** (matches local exactly) |
| `MBW Downloadable re-branding\_rebrand-templates` | Kit artwork present — Portrait, Landscape, `_psd-masters`, `_preview.html` |
| `MBW Downloadable re-branding\_delivery-manifests` | `README.txt`, `filename-aliases.csv`, `manifests` |

Also on Drive in that same folder: **`Budget Mailboxes Downloadable Asset Rebranding SOP.gdoc`** — the SOP as a
Google Doc, alongside the `designer-sidekick` path the playbook cites.

**Access is not a blocker.** Kunchana Godahewa shares the Drive link with each new batch, so a successor picks
up access the next time a batch is handed over, and **the brand kit is built per batch anyway** — it is not a
standing asset that has to be inherited.

### The trap when you pull a kit from Drive

**The MBW kit on Drive has no `brand.json`** — only the artwork folders. So does the next kit Kunchana shares,
most likely. A kit with artwork but no `brand.json` **does not fail**: `BrandKit` silently falls back to
defaults, which means zero manufacturer aliases, no tagline, no disclaimer, and the brand name resolves to
*Budget Mailboxes* — on a MailboxWorks job, or any other. It looks like it worked and prints raw,
inconsistent manufacturer names.

The approved naming decisions for the two delivered brands are in this repo at **`brandkits/`**. When you
build a kit for a new batch, author its `brand.json` — starting from `docrefine/assets/brand.example.json` —
and **save it beside the artwork under that literal filename**. Any other name is ignored.

## 6. Open ClickUp work assigned to Jason

- **[86eypep2z](https://app.clickup.com/t/86eypep2z) — Sustainable Hearth PDFs.** Assigned to him, scope not
  agreed. See [02-sustainable-hearth.md](02-sustainable-hearth.md). Needs reassigning or closing.
- **[86ex00r23](https://app.clickup.com/t/86ex00r23) — the DocRefine Pro project task.** How this project is
  reported upward. Needs a new owner or it drifts, which has happened before (it once fell four versions behind).

## 7. Machine-specific setup that has to be rebuilt

Nothing here is hard, but none of it is in the repo:

- Python 3.10.11 at a specific path — the `python`/`py` commands on PATH were Microsoft Store stubs and must
  not be used. See [knowledge/project-setup.md](knowledge/project-setup.md).
- **Ollama plus local models** for classification and the vision pass. The vision pass was measured on an
  **RTX 5060 with 8GB VRAM**, and that limit shapes the design: two models resident at once drops throughput
  from 3.5s to 14.3s per file. A successor on weaker hardware should read
  [knowledge/vision-pass.md](knowledge/vision-pass.md) before promising a timeline.
- Local git identity was `Jason Diaz <jason@weblifestores.com>` — change it.

## 8. What is NOT in this repo, deliberately

- **Any API token or credential.** Jason's ClickUp token lives in a personal Claude Code skill
  (`personal-config`) that is explicitly marked non-shareable. It is not here and must not be added.
- **His personal ClickUp context** — 1:1 DM channel IDs and private notes about colleagues, from a
  `personal-context` skill. Not business-critical and not his to hand over.
- The `weblife-clickup-config` / `weblife-clickup-ops` / `clickup-rich-comment` skills are **workspace-wide**
  and already shared with the team — a successor in Task Agency should already have them. They are what make
  ClickUp comments with real @mentions possible; the plain MCP tool cannot do it. If a successor does not have
  them, ask Task Agency, not this repo.
