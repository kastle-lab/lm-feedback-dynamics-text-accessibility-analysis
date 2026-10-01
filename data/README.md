# Data

This directory contains the source chapter excerpts used as inputs for the text-accessibility rewriting experiments.

## Files

File  | Intent
--- | ---
`chapters.json`  | Stores the original educational chapter excerpts and metadata used by the rewriting pipeline.

## Record Structure
Field  | Description
--- | ---
`no.`  | Chapter-excerpt identifier used to track the item across results.
`book`  | Source textbook or volume name.
`pages`  | Page range for the source excerpt.
`header`  | Chapter or excerpt heading.
`text`  | Original chapter text passed to the rewriting pipeline.