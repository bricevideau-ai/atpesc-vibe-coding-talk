# Vibe Coding from a System Software and Performance Engineering Perspective

> ATPESC talk · 20 minutes · **20 content slides (25 total) — at the hard ceiling.** A
> retrospective on AI-agent-assisted development across two deliberately contrasting systems
> projects. Anything added from here must displace something.
>
> Authors: Brice Videau (ALCF) and Claude. The deck is itself a vibe-coding artifact —
> pretending otherwise would undercut the talk.

## Thesis

Two claims, and the talk fails if it lands only one of them.

**The gain is real and large.** Six weeks on CCS produced roughly two years of work at the prior
rate — 96 merged PRs on a C library with bindings. And rust-gpu/claspr is a gain of a different
kind entirely: a SPIR-V backend port needing deep Rust and SPIR-V expertise Brice does not have.
Not work that would have gone slowly — work that would not have happened. *Compression* on one
project, *reach* on the other.

**And you cannot let it run blindly.** Every single defect in this corpus was caught by **a
sanitizer, a tracer, a bisect, a second implementation, or a human reviewer** — never by the agent
re-reading its own code.

The two case studies differ on one axis: **can the human verify the work?**

| | CCS | rust-gpu + claspr |
|---|---|---|
| Domain | Online autotuning middleware for HPC runtimes (ECP continuation), production C99 | rust-gpu from graphics SPIR-V to OpenCL SPIR-V, targeting Aurora; grew into a single-source Rust GPU programming model |
| My expertise | Complete — I wrote it and maintain it | Adjacent only — OpenCL expert, not Rust or SPIR-V |
| Agent's footing | Ordinary, well-represented C | Uncharted; nothing comparable in the training corpus |
| Verification | Direct — read the diff | Externalized — differential runtimes, difftests, bisect, tracers, golden oracles |
| My role | Reviewer and taste-keeper | Oracle designer and spec authority |

## Contents

| File | What it is |
|---|---|
| `build_deck.py` | `python-pptx` generator. **Edit this, not the pptx.** |
| `ATPESC 2026 - Vibe Coding - Brice Videau.pptx` | The deck. |
| `ATPESC 2026 - Vibe Coding - Brice Videau.pdf` | PDF render for review. |
| `ALCF Presentation Template.pptx` | Committed so the build is self-contained. |

## Rebuild

Filenames contain spaces — quote them.

```bash
DECK="ATPESC 2026 - Vibe Coding - Brice Videau"
python3 build_deck.py                                # asserts slide count == 25
soffice --headless --convert-to pdf "$DECK.pptx"
rm -f slide-*.jpg && pdftoppm -jpeg -r 90 "$DECK.pdf" slide
```

The `slide-*.jpg` files are ephemeral QA artifacts — regenerate, don't commit.

## Design rule: sparse slides, rich notes

**The detail lives in the speaker notes, not on the slide.** 21 of 25 slides carry notes with the
quotes, exact figures, commit SHAs and asides. On-slide bullets stay at ~4 per column.

This is not a style preference — it was forced by measurement. The first build put the detail on
the slides and a visual QA pass found **13 of 18 content slides overflowing**, several clipped
mid-word over the DOE footer. Layout geometry that bites, measured from the template:

| Placeholder | Reality |
|---|---|
| idx 13 "subtitle", all content layouts | 12.22" × **0.64"**, `noAutofit` — one line, ~80 chars |
| Layout 4 | Has **no body placeholder**; idx 13 is all you get, so it cannot hold prose. Use Layout 3. |
| Layout 3 body idx 14 | 10.67" × 4.77", ends at y=6.86"; the footer begins immediately after |
| Layout 9 block header | Must fit **one line** (~20 chars) or it eats body height and clips the body |
| `set_code` column | 5.96" × 3.97" at 9pt Consolas; hard-wrap your own lines |

Re-run the visual QA (a subagent that *reads each JPG as an image*) after any content change. A
byte or file-size check is not a substitute — that mistake has its own memory entry.

## Relationship to the ALCF talk

`../alcf-ai-agent-coding-talk` ("From Directing to Dialogue", 38 slides, June 2026) covers the same
two projects for a different audience and at nearly twice the length. This deck is **not** a
retarget: different thesis (cautionary rather than evaluative), a third of the length, and it adds
June–July work the ALCF deck predates — the four merged upstream pocl PRs, the command-buffer race
hunt, and the context-as-a-resource refactor.

The helper block in `build_deck.py` (`remove_all_slides` … `set_code`) is ported **verbatim** from
that deck. It encodes five already-diagnosed rendering bugs; don't "simplify" it.

## Provenance of every number on a slide

All re-derived from source on **2026-08-05**. An earlier mining pass got the CCS PR count wrong by
sampling only #78–#122; the figures below are from the full query. **If you change a number,
re-derive it — do not copy it from the ALCF deck.**

Slide numbers below are for the **25-slide build**. Regenerate the index after any insertion:

```bash
python3 -c "
from pptx import Presentation
p=Presentation('ATPESC 2026 - Vibe Coding - Brice Videau.pptx')
for i,s in enumerate(p.slides,1):
    t=next((ph.text_frame.text.replace(chr(10),' / ') for ph in s.placeholders
            if ph.placeholder_format.idx==0), '(closing)')
    print(f'{i:2d}  {t}')"
```

| Number | Slide | Source |
|---|---|---|
| 99 PRs, #24→#122, 96 merged / 3 closed | 9, 10 | `gh pr list -R argonne-lcf/CCS --author bricevideau-ai --state all --limit 300` |
| Feb 26 → Apr 7 (41 days), ~16 merged PRs/week | 9 | derived from the same query's `createdAt`/`mergedAt` |
| "Six weeks ≈ two years at the rate I had been going" | 10, 24 | **Brice's own estimate as the maintainer, not a measurement.** Flagged as such in the speaker notes so it can be defended in Q&A. |
| ~53,000 lines of C | 6 | `wc -l` over `src/` + `include/` + bindings on `devel` |
| rust-gpu +24,262 / −384 across 765 files | 9 | `gh pr view 3 -R bricevideau-ai/rust-gpu --json additions,deletions,changedFiles` |
| OpenCL **1.2 and 2.0** target environments | 9 | `.github/workflows/ci.yaml:176` — compiletests run `--target-env …,opencl1.2,opencl2.0`. **Corrected:** an earlier draft claimed "8 `spirv-unknown-opencl*` targets", derived by grepping strings. That grep hit an exhaustive `match` over spirv-tools' pre-existing `TargetEnv` enum in `link.rs` (the same block lists OpenGL 4.0–4.5 and WebGPU). Two is the real number. |
| claspr 466 commits, ~49K Rust LOC, 417 tests, 24 compile-fail fixtures | 9 | `git log --oneline main \| wc -l`; `find -name '*.rs' -not -path './target/*' \| xargs wc -l`; `grep -c '#\[test\]'`; `find -path '*compile_fail*' -name '*.rs'` |
| 4 pocl PRs merged (#2166, #2214, #2215, #2216); 2 issues (#2174, #2175) | 9, 12 | `gh pr list -R pocl/pocl --author bricevideau-ai --state all` |
| #2214 failure mode and `-14` outcome | 12 | pocl commit `b55a569a5` message |
| 7 review comments, "contradicts the comment in pocl_cl.h" | 12 | pocl PR #2214 review thread (maintainer `jansol`) |
| Fixes in Mesa/rusticl and Intel compute runtime | 12 | Worked directly with Karol Herbst (Red Hat/Mesa), who filed [Mesa MR !41404](https://gitlab.freedesktop.org/mesa/mesa/-/merge_requests/41404), and Ben Ashbaugh (Intel). Not filed under `bricevideau-ai`, so not discoverable via `gh`. |
| "host wrote 6, kernel sees 0" | 20 | memory `project_hostbuffer_coherency_rusticl.md` |
| bisect green `a049ad4` / first red `236d9c0`; `wait=NONE` → `wait=[3]` | 21 | claspr commits `89386bf`, `3995ccb` |
| "~200K tokens before it can make ANY change" | 22 | `claspr/NOTES.md:79` |
| cognitive 39→7; cyclomatic 66→23, cognitive 33→13 | 22 | `claspr/NOTES.md:213–235`, Mozilla `rust-code-analysis-cli` |
| 11,023 → 3,674 lines (−67%) | 22 | `claspr/NOTES.md:98` |
| 1,000 Scientist AI Jam Session, **February 28, 2025** | 2, 3 | [anl.gov/cels/1000-scientist-ai-jam-session](https://www.anl.gov/cels/1000-scientist-ai-jam-session) and [events.cels.anl.gov/event/611](https://events.cels.anl.gov/event/611/). Note: Brice first recalled 2024; the model lineup (Claude 3.7 Sonnet shipped four days before the event) and the Indico page both confirm **2025**. |
| The four models, their failures, and the ranking | 2, 3 | Brice's own contemporaneous written assessment, quoted in the speaker notes |
| 1,000 Scientist AI Jam Session, **February 28, 2025** | 2, 3 | [anl.gov/cels/1000-scientist-ai-jam-session](https://www.anl.gov/cels/1000-scientist-ai-jam-session) and [events.cels.anl.gov/event/611](https://events.cels.anl.gov/event/611/). Note: Brice first recalled 2024; the model lineup (Claude 3.7 Sonnet shipped four days before the event) and the Indico page both confirm **2025**. |
| The four models, their failures, and the ranking | 2, 3 | Brice's own contemporaneous written assessment, quoted in the speaker notes |

Deliberately **excluded**: token and dollar figures. The only snapshot available is frozen at
2026-06-09, covers 5 of 17 sessions, and `session_stats.json`'s token volumes disagree with
`ccusage` by roughly 2×. Not solid enough to project on a wall.

## The opening, and why it works

Slides 2–3 open on the **1,000 Scientist AI Jam Session (Feb 28, 2025)**, where Brice submitted
*this exact port* to four frontier reasoning models and graded the results in writing at the time.
That assessment is the control condition for the whole talk: one model documented, in the README
and developer guide, changes to a backend it had never touched.

It also closes a loop. In 2025 Brice partly blamed the failure on rust-gpu's backend being
"a plate of spaghetti." Slide 22 returns to that: a year later he stopped treating illegibility as
an excuse and started measuring it with `rust-code-analysis`.

## Corrections applied after review

Worth recording, because three of these were confident errors of exactly the kind the talk warns
about:

- **"Maintainers had no idea an agent was involved"** — false, and I wrote it while holding the
  contradicting evidence. `bricevideau-ai` is transparently an agent account, every commit is
  co-signed by Claude Code, and Brice wrote *"Claude's working on it"* in the pocl PR thread. Slide
  12's notes now say the opposite, and it makes a better story: upstream maintainers knowingly
  reviewing agent-authored patches.
- **"8 OpenCL targets"** — wrong, see the provenance table. Two.
- **Apple OpenCL** — I asserted it behaves permissively like pocl. We never tried it. Removed.
- **Missing ecosystem impact** — the work also produced fixes in Mesa/rusticl and Intel's compute
  runtime, not just pocl. Slide 12 now covers all three.
- **Missing multi-model review** — a second model (ChatGPT 5.5) reviewing claspr's implementation
  drove substantial changes. Now slide 4 and the "second opinion" gate on slide 23.
- **Undersold the upside** — the deck proved the caution and only implied the gain, risking a
  "don't bother" takeaway. Slide 10 now states it directly, and the takeaway leads with it.
