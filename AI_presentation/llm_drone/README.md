# LLM-guided drone path planning

16:9 LaTeX Beamer discussion of *Goal-Oriented Intelligence Path Planning for LLMs Based Drone in Adversarial Conditions*, by Muhammad Sohail Farooq et al. Presenter: **Arush Ram Anandh** (reflecting the local author edit).

The deck follows `codex_slide_plan.md`: a separate title page and eight substantive slides, with approximately 18.5 minutes of presenter notes embedded in `main.tex`. `paper.pdf` supplies the technical evidence; `old.pdf` supplies visual references only. Reported evidence, missing details and proposed extensions are explicitly distinguished.

## Build

Requires a TeX Live installation with Beamer, Latin Modern, TikZ/PGFPlots, standalone, booktabs and latexmk. From this directory:

```sh
latexmk -pdf -jobname=presentation -interaction=nonstopmode -halt-on-error main.tex
```

This produces `presentation.pdf`. The vector figure crops are already supplied under `assets/`; no shell escape, network access or external fonts are required.

To rebuild the crops from the original paper:

```sh
cd assets
pdflatex -interaction=nonstopmode -halt-on-error figure3.tex
pdflatex -interaction=nonstopmode -halt-on-error figure4.tex
cd ..
```

The crop sources retain the original PDF content, including diagram labels and all four Figure 4 values. Table I is editable LaTeX; the motivation schematic, line icons and flow diagrams are editable TikZ.

To create a separate presenter-notes PDF:

```sh
pdflatex -jobname=presenter-notes -interaction=nonstopmode -halt-on-error '\def\shownotes{1}\input{main.tex}'
pdflatex -jobname=presenter-notes -interaction=nonstopmode -halt-on-error '\def\shownotes{1}\input{main.tex}'
```

## Evidence and interpretation

The one-page audience handout is `reference_sheet.pdf`, with editable source in `reference_sheet.tex`. Build it with:

```sh
latexmk -pdf -interaction=nonstopmode -halt-on-error reference_sheet.tex
```

- The paper interleaves remote few-shot prompting and fine-tuning without identifying the system behind its aggregate metrics.
- The clearance values conflict: 2 m on p. 4 and 1.2 m on p. 5.
- The reported results concern simulation. Trial counts, uncertainty, latency and controlled planner baselines are absent.
- Figure 4 labels the success percentages as “efficiency”; the paper defines no separate path-efficiency measure.
- Ordinary RRT does not guarantee shortest paths. Interpolation alone guarantees neither safety nor dynamically feasible flight.
- Slide 7 is explicitly our analysis and a proposed extension, not an evaluated paper contribution.

Paper page numbers refer to the seven-page supplied PDF. Source notes appear on every substantive slide. The timing and methodological qualifications appear in the embedded presenter notes.
