# pptxjs — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Open

_None tracked yet._

## Done

- Chart and SmartArt graphic-frame fallback rendering (`feat/p8-chart-smartart-fallback`). Charts emit their cached preview image when present, else a `[Chart]` placeholder at the frame position. SmartArt frames expand the sibling `diagrams/drawingN.xml` DrawingML cache through the standard shape pipeline (offset by the frame origin), falling back to a `[SmartArt]` placeholder when no drawing cache exists.
