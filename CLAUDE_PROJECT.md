# CLAUDE.md — Spreadsheet Application Project

## Project Vision
We are building a spreadsheet application that beats Microsoft Excel. Our goal is to create the most beautiful, performant, and feature-complete spreadsheet the world has never experienced. Web-first, cross-platform, AI-native.

## Critical Reference File
**ALWAYS consult `EXCEL_FEATURES.md` before implementing any spreadsheet feature.**
This file contains 1000+ individual features mapped from Excel with priority levels, customization depth, and status tracking.

## Rules for Feature Implementation

### 1. Never Build a Feature Shallow
When implementing ANY spreadsheet feature, check `EXCEL_FEATURES.md` for the FULL depth of that feature. For example:
- "Add number formatting" → check ALL number format types, custom format codes, 4-section format strings
- "Add borders" → check all 13 border styles, per-edge control, diagonal borders, border colors
- "Add chart" → check ALL customization options (axes, trendlines, data labels, error bars, etc.)

### 2. Update Status in EXCEL_FEATURES.md
After implementing a feature, update its status:
- ⬜ → 🟡 (in progress)
- 🟡 → ✅ (done and tested)
- ⬜ → ⏭️ (explicitly deferred with reason)

### 3. Priority Order
Build in this order:
1. P1 items in the current focus area
2. P1 items in dependent areas
3. P2 items in completed areas
4. P3 items only when all P1+P2 in that area are done

### 4. .xlsx Compatibility is Sacred
Every feature MUST round-trip through .xlsx without data loss. If a feature can't be saved to .xlsx, it must gracefully degrade (preserve data, lose only unsupported customization).

### 5. Formula Engine is the Foundation
The calculation engine (dependency graph, recalculation, spill engine) must be rock solid before adding UI features that depend on it. Get the compute layer right first.

## Architecture Principles
- **Web-first**: The web version is the FULL version, not a companion
- **Computation off main thread**: Calc engine in Web Worker (Rust/WASM preferred)
- **Canvas/WebGL rendering**: No DOM-per-cell approach
- **CRDT for collaboration**: Not OT — CRDTs handle offline + real-time better
- **AI-native**: AI woven into every layer, not bolted on

## Code Quality
- Every formula function must have unit tests matching Excel's behavior
- Grid rendering must maintain 60fps during scroll at 100K+ rows
- .xlsx import/export must be tested against real-world Excel files

## File Structure Reference
- `EXCEL_FEATURES.md` — Master feature checklist (1000+ items)
- `state.md` — Current sprint progress and blockers (update frequently)
- `architecture.md` — Technical architecture decisions and diagrams
