# pptxjs — TODO

Tracked work for this fork. Move entries into the "Done" section below as they ship; link the PR / commit.

## Open

_None tracked yet._

## Done

- Hyperlink rendering (P9, Wave 2): external URLs and intra-deck slide jumps
  on text runs, shapes, and images are wrapped in `<a>` anchors. External
  links use `target="_blank" rel="noopener noreferrer"`; intra-deck jumps
  (ppaction hlinkshowjump + slide-to-slide rels) become `#slide-N` fragments.
  URL scheme whitelist (http/https/mailto/tel + fragments/relatives) rejects
  `javascript:` and other unsafe targets.
