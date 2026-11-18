# Data-quality issues, assumptions and decisions

Numbers come from `results/eda/trip_stats.csv` and `data_quality_summary.md` (scripts/01_eda.py).
Partition used: train (80 trips) unless stated; student_test (10 trips) was only inventoried/sanity-checked.

| # | Observation | Decision |
|---|---|---|
| 1 | Every trip ships its **own** road network (299 – 155k nodes, 649 – 307k arcs) in a different city/country. | Process each trip independently; never use absolute lon/lat as model features (would not generalise). |
| 2 | Coordinates are lon/lat degrees. | Per-trip local equirectangular projection to metres; max single-step error vs haversine 0.72 m (train) — negligible vs GPS noise. |
| 3 | `.route` has far fewer rows than GPS points (≈5 points per route arc). | There are no per-point labels. We derive them by a monotone alignment of GPS points to the ordered route (Phase 2). |
| 4 | Reference route is contiguous (contiguity ≥ 0.98) and has no invalid arc ids. | Route order can be trusted for the monotone alignment. |
| 5 | Some routes cover only part of the track (train 00000033: 6.4 km route vs ~48 km of track, 67 % of points > 40 m away; 00000054: 88 %; 00000002: 13 %; 00000043: 9 %). | Points > 40 m from the reference route get label −1 ("unmatched"); excluded from training positives and from the main accuracy denominator; strict accuracy (counting them as errors) is reported too. |
| 6 | Sampling nominally 1 Hz but gaps exist: 28 train trips have a gap > 30 s, the largest is 153 064 s (~42 h). | HMM breaks the chain at large gaps (dt > threshold) and restarts; time-gap features are given to the classifier. |
| 7 | 147 implied-speed outliers (> 70 m/s) in 10 train trips. | Kept (no labels are dropped), but kinematic features are computed from smoothed displacement and clipped. |
| 8 | Arcs are directed; ~95 % have a reverse twin with identical geometry. | Direction of travel (heading vs arc bearing) is a key feature; the HMM penalises U-turns. |
| 9 | Networks contain self-loops, duplicate (parallel) arcs and zero-length arcs. | Shortest paths use the shortest parallel arc; self-loops ignored in routing; zero-length arcs get a tiny positive weight. |
| 10 | Networks are not fully strongly connected (largest SCC ≥ 63.6 % of nodes; up to 421 weak components). | Unreachable candidate transitions get probability 0; if no transition is feasible the HMM chain is split instead of failing. |
| 11 | No missing / non-finite values, no duplicate timestamps, all tracks already time-sorted, all tracks inside their network bounding box. | Cleaning code is kept (sort, drop non-finite, drop duplicate timestamps) for robustness on unseen evaluator trips. |
| 12 | Reference routes are reference information, not verified ground truth (dataset doc §6). | Derived labels inherit any route errors; we report this as a limitation and avoid over-claiming accuracy. |
