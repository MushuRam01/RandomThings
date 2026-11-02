# Codex Implementation Plan — LLM-Guided Drone Path Planning

## Objective
Build a technically accurate, visually restrained **16:9 LaTeX Beamer** research presentation critically discussing *Goal-Oriented Intelligence Path Planning for LLMs Based Drone in Adversarial Conditions* by Muhammad Sohail Farooq et al.

- **Presenter:** Arush Ramanand (`\author{Arush Ramanand}`)
- **Primary source:** `paper.pdf` (seven-page paper)
- **Visual reference only:** `old.pdf` (older presentation; contains unsupported or unnecessary material)
- **Deliverables:** `main.tex`, `assets/`, compiled `presentation.pdf`, and `README.md` with build instructions.
- **Target duration:** 15–20 minutes; eight substantive slides. A title slide may be added separately if desired, but preserve the slide sequence below.

## Non-negotiable accuracy rules

1. `paper.pdf` is authoritative for what the authors **state or report**. `old.pdf` is **not** an authoritative technical source.
2. Separate **reported**, **inferred/reconstructed**, and **not specified**. Do not present a plausible implementation as a documented one.
3. **RRT is a classical randomized motion-planning algorithm, not machine learning.** Ordinary RRT does not guarantee shortest paths; do not repeat the paper's claim that it does.
4. The paper discusses both **few-shot prompting of commercial LLMs/VLMs** and **fine-tuning a pretrained LLM**. It does not adequately establish their relationship or which system generated the aggregate reported results.
5. The paper reports **2 m clearance** on page 4 but **1.2 m** on page 5. Flag the discrepancy rather than resolving it.
6. Do not invent the fine-tuned model's identity, dataset size, train/test split, number of trials, obstacle velocities, latency, replanning rate, hardware configuration, exact prompt template, coordinate tokenization, or controller type.
7. The paper's results are from **simulated environments**. Claims of real-world, GPS-denied, or bad-weather validation are not supported by the presented evaluation.
8. Do not claim interpolation ensures smoothness, obstacle avoidance, or dynamically feasible flight. It only adds points along existing segments.
9. Any added mathematical formulation must be explicitly marked **explanatory reconstruction**, not an equation supplied by the authors.
10. Include unobtrusive paper page/figure/table references on substantive slides.

## Visual style

Use a white or near-white background with a coherent **RGB purple** theme, not the teal palette in `old.pdf`.

| Role | Hex | RGB |
|---|---|---|
| Primary purple | `#673AB7` | `(103,58,183)` |
| Deep purple | `#4527A0` | `(69,39,160)` |
| Light purple | `#B39DDB` | `(179,157,219)` |
| Pale lavender | `#F3EEFA` | `(243,238,250)` |
| Body text | `#242136` | `(36,33,54)` |

- Beamer, 16:9; Metropolis or a similarly restrained custom theme.
- Dark body text, purple titles, thin purple separators, sparse pale-lavender emphasis.
- Use `tikz`, `pgfplots` (if needed), `booktabs`, `amsmath`, `graphicx`, `hyperref`.
- Consistent typography, generous margins, legible diagrams, and compact source footers.
- No stock imagery, decorative gradients, excessive boxed callouts, tiny text, or needless equations.
- Preserve the **small drone/obstacle/goal illustration** idea from page 2 of `old.pdf`, but recreate it as a clean TikZ graphic.
- Preserve the **original Figure 3 architecture image** treatment from page 5 of `old.pdf`, but extract a clean crop from `paper.pdf` page 4.
- The **component comparison table** from page 3 of `old.pdf` should **not** appear on the opening motivation slide. Move a concise comparative table to Slide 7.

---

## Slide 1 — Motivation: Can Foundation Models Help Drones Navigate?

**Purpose:** Introduce the problem and the research question, not implementation details.

**Layout:** Approximately 55% text on the left, 45% minimalist TikZ illustration on the right, inspired by `old.pdf` page 2. Show a drone/start, obstacles, goal, and a proposed route. Do not imply the illustration is a measured result.

**On-slide content:**
- Navigate from a starting position to a destination while avoiding obstacles.
- Unfamiliar or changing environments motivate adaptive navigation.
- The paper explores **LLMs/VLMs for visual-context-guided waypoint planning**.
- Central question, emphasized: **What useful role can a foundation model play in a drone navigation system?**

**Presenter notes:** Explain the motivation and the distinction between conventional geometric planners and learned foundation-model-based decision-making. Do not characterize RRT as ML or assert that the paper establishes superior navigation.

**Source:** `paper.pdf` pp. 1, 3–4.

## Slide 2 — Why LLMs? The Authors' Broader Motivation

**Purpose:** Explain the paper's ambitious motivation while distinguishing it from the actual experiments.

**Layout:** Three horizontally arranged sections, each with a simple line icon and at most two lines of explanatory text.

1. **Coordination:** Multi-UAV communication, mission planning, information exchange.
2. **Contextual reasoning:** Interpret instructions, heterogeneous observations, unexpected events, and anomalies.
3. **Adaptation:** Respond to changing obstacles and generate goal-directed navigation decisions.

**Bottom statement:** **Specific experiment:** vision-guided waypoint generation in simulated obstacle environments.

**Presenter notes:** The authors discuss broader coordination, communication and emergency-response capabilities; these are motivations, **not experimentally demonstrated contributions** of this paper. The experiment is narrower.

**Source:** `paper.pdf` p. 3.

## Slide 3 — Proposed Architecture: From Observations to Flight

**Purpose:** Walk through the actual system diagram.

**Layout:** Use **Figure 3 from `paper.pdf` page 4**, cropped to the diagram, occupying ~70–75% of the slide. On the right, show a short numbered reading guide:

1. Start/goal/instructions + camera observations.
2. YOLOv8 detects objects and provides visual context.
3. Remote LLM/VLM receives the constructed input.
4. Waypoint response is parsed and interpolated.
5. Trajectory is passed through ROS to the motion controller.

**Small caveat:** **Exact fine-tuned versus remotely prompted model arrangement: not specified.**

**Presenter notes:** YOLOv8 is a learned object detector, not by itself a 3D coordinate estimator. RGB and depth observations are mentioned, but conversion to obstacle world coordinates is not explained. The paper suggests updated observations may allow replanning, but does not specify replanning cadence or trigger. Interpolation is approximately 0.1 m between points; it does not by itself guarantee smoothness or safety. PID is a plausible low-level controller but **not explicitly identified** in the paper.

**Source:** `paper.pdf` p. 4, Fig. 3 and methodology; p. 5 for RGB/depth.

## Slide 4 — Fine-Tuning, Prompt Structure, and Inference

**Purpose:** Untangle the paper's interleaving of dataset generation, model training, and deployment.

**Layout:** Two parallel flow diagrams, **Training** on left and **Inference** on right, followed by one equation and a small methodological-gap strip.

**Training column:**

`Simulated start + goal + obstacles` → `RRT-generated reference waypoints` → `Fine-tune pretrained LLM (cross-entropy)`.

**Inference column:**

`Start + goal + current visual context + instruction` → `LLM/VLM` → `Waypoint sequence` → `Parse/interpolate/control`.

**Show the paper's input expression:**

\[
\mathrm{Input}=\mathrm{Start+Goal+Visual\ Map+Instruction}.
\]

**Reported:** RRT-generated waypoint targets; 105 epochs; cross-entropy loss; reported training loss around 0.15 near epoch 65; three example visual maps for few-shot prompting.

**Not specified:** Exact model and fine-tuning method, dataset size, split, tokenization/coordinate formatting, exact prompt, visual-to-spatial transformation, and whether reported metrics come from the fine-tuned or prompted models.

**Presenter notes:** Explain RRT as a randomized collision-free tree search. It generates reference trajectories, while the LLM apparently learns to imitate waypoint outputs. The paper does not clearly establish whether training used structured coordinates, images, or both. Avoid claiming a documented token-level loss equation or a verified teacher–student implementation beyond the stated evidence. Mention the **2 m versus 1.2 m** clearance inconsistency if time permits.

**Source:** `paper.pdf` pp. 4–5.

## Slide 5 — Experimental Results: What Was Measured?

**Purpose:** Display the actual quantitative evidence, with Table I and Figure 4 **both on this slide**.

**Layout:** Left: recreate **Table I** in LaTeX using `booktabs`. Right: embed a clean crop of **Figure 4 from `paper.pdf` p. 5**. Under both, a short evidence/limitations strip.

| Environment | Obstacles | Moving obstacles | Success rate | Average completion time |
|---|---:|---:|---:|---:|
| Simple | 8–10 | 3 | 95.0% | 65.7 s |
| Moderate | 10–12 | 7 | 92.0% | 65.4 s |
| Complex | 12–16 | 10 | 91.5% | 69.2 s |
| Dynamic | 16–20 | 12 | 89.0% | 72.0 s |

**Short notes beneath:**
- **Simple/moderate/complex/dynamic:** obstacle counts are reported, but **formal difficulty definitions are not specified**. The dynamic scenario includes moving obstacles; their motion model is not specified. All four categories list moving obstacles.
- **Number of simulation trials/runs:** **Not specified**.
- Figure 4 labels the success percentages **“efficiency”**, without a separate defined path-efficiency metric.
- Success decreases **95% → 89%** across the four reported categories; this is descriptive, not proof of a particular statistical relationship.

**Presenter notes:** Interpret the table cautiously. No reported standard deviations, confidence intervals, explicit trial counts, direct RRT baseline, or quantified path optimality. The paper's real-world and bad-weather assertions are not demonstrated by the displayed simulation results.

**Source:** `paper.pdf` p. 5, Table I and Fig. 4; p. 6 conclusion.

## Slide 6 — What Is Missing from the Evaluation?

**Purpose:** Give a compact, evidence-based methodological critique.

**Layout:** A two-column table, styled similarly to the compact comparison table in `old.pdf` page 3, but purple and more concise.

| Dimension | Missing or unclear |
|---|---|
| Reproducibility | Model identity, dataset size, trials, splits |
| Baselines | Direct RRT/RRT*, A*, or other controlled comparisons |
| Safety | Collision rates, segment validation, measured clearance |
| Path quality | Length, optimality, smoothness, energy |
| Real-time operation | Inference latency, replanning frequency, compute |
| Generalization | Unseen environments, sensor noise, physical flight |

**Bottom research question:** **What advantage does the learned planner demonstrate over the planner used to generate its reference paths?**

**Presenter notes:** Distinguish missing evidence from evidence of failure. The architecture may be interesting, but the results do not isolate an LLM benefit or validate real-world reliability.

**Source:** `paper.pdf` pp. 4–6.

## Slide 7 — Alternative Planning Methods and Possible Extensions

**Purpose:** Show what the paper could have compared, and propose a reasonable research direction. Label this slide **Our analysis / proposed extensions**.

**Layout:** Compact four-row comparison table; underneath, a three-stage hybrid pipeline.

| Approach | Role | Why include it? |
|---|---|---|
| A*, RRT, RRT* | Classical geometric planning | Feasibility, path cost, reproducible baseline |
| Model Predictive Control (MPC) | Receding-horizon planning/control | Dynamic obstacles and control constraints |
| Reinforcement Learning (RL) | Learned navigation policies | Learning from interaction and adaptation |
| Hybrid LLM + classical planner | Semantic high-level planning + geometric validation | Test added value without delegating safety entirely to generated tokens |

**Proposed pipeline:**

`LLM/VLM: mission interpretation and high-level decisions` → `Conventional planner: generate and validate feasible route` → `Low-level controller: track trajectory`.

**Presenter notes:** This is **not** an architecture the paper has evaluated. Discuss why an LLM might be better suited to interpreting instructions, changing goals, or semantic context than directly replacing a proven collision checker. Keep the comparison fair; do not claim RL or MPC necessarily outperforms the paper without experiments.

## Slide 8 — Conclusion: Promise Versus Evidence

**Purpose:** Close with a balanced, defensible assessment.

**Layout:** Two short columns, then one prominent final sentence.

**What the paper offers:**
- An integration concept connecting visual perception, LLM/VLM planning, waypoint processing, and drone control.
- Preliminary simulated goal-reaching rates across four obstacle categories.

**What it does not establish:**
- Reproducible fine-tuning and inference details.
- Comparative benefit over established planners.
- Safety, real-time feasibility, and real-world generalization.

**Final sentence:** **Foundation models may add value at the high-level decision layer, but that value needs controlled evaluation against conventional planning methods.**

**Presenter notes:** Frame the paper as a motivation for research rather than a conclusive demonstration of superior autonomous flight.

---

## Figure extraction and asset instructions

- Extract **Figure 3** from page 4 of `paper.pdf` (architecture). Crop out the surrounding two-column body text; keep all blocks, arrows, and labels visible.
- Extract **Figure 4** from page 5 of `paper.pdf` (success/“efficiency” chart). Preserve the original plot and its four values.
- Recreate **Table I** in LaTeX, rather than pasting the low-resolution image.
- Recreate the small **drone / obstacle / goal** drawing from `old.pdf` page 2 using TikZ. It is an illustrative schematic, not experimental data.
- The table on `old.pdf` page 3 is a **stylistic reference** for Slide 7, not a source of factual claims.
- `old.pdf` page 5 is a useful reference for **Figure 3 placement**, but use a fresh crop from `paper.pdf` rather than copying the surrounding paper text.
- Do **not** reproduce the unsupported mathematical reconstructions on `old.pdf` pages 6–8 unless independently justified and explicitly labeled as reconstructions.

## Build and QA checklist

1. Write a self-contained `main.tex` with reusable colors, title/footer macros, and figure/table styles.
2. Store all figures under `assets/` and use relative paths.
3. Compile using `latexmk -pdf main.tex` (or document the available equivalent).
4. Render all PDF pages and visually inspect **every slide** for overflow, cropping, alignment, and text size.
5. Verify **all four Table I rows** against `paper.pdf` and that Figure 4 matches the paper.
6. Verify presenter name **Arush Ramanand** on the title/author metadata and no inherited incorrect author spelling from `old.pdf`.
7. Confirm all source notes identify `paper.pdf` pages accurately.
8. Verify no unsupported statements have been introduced about training details, flight hardware, latency, guarantees, real-world tests, or superiority.
9. Deliver `main.tex`, `assets/`, `presentation.pdf`, and `README.md`.

## Suggested speaking times

| Slide | Minutes |
|---|---:|
| 1 — Motivation | 1.5 |
| 2 — Why LLMs? | 2.0 |
| 3 — Architecture | 3.0 |
| 4 — Fine-tuning and inference | 3.5 |
| 5 — Results | 3.0 |
| 6 — Evaluation gaps | 2.0 |
| 7 — Alternatives | 2.5 |
| 8 — Conclusion | 1.0 |
| **Total** | **18.5** |
