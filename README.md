# WJEC Item Level Analysis Cleaner
A Google App Script that takes the standard WJEC Item Level Analysis data they provide in their portal and reorganises the data so you can sensibly analyse it.

By default, the item level analysis data from the WJEC portal is awkwardly formatted so that the `Max Mark`, `Question Number` and `Mark` columns are all separate, making it really difficult to any sensibly analysis. This script creates two headers, with the top row being `Question Number` and beneath it being `Max Mark` for each question. The marks each student gets are then put underneath making it easy to analyse the data afterwards. It also conditionally formats the item level data, using the 'Max mark' value as the Max point. 

This is mostly put here because I only do this once a year and means I hopefully won't end up re-inventing the wheel. Hopefully it's useful to others.

If it is, and you want to extend further, please open a PR. Contributions are most welcome!
