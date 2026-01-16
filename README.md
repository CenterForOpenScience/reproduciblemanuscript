This is a proof of concept code intended to demonstrate an R-based workflow for reproducible manuscripts. This code is for demonstration only, and will not be maintained.

The main function in this package is the knit_docx() function. This function knits together an output .docx document from a template .docx document with placeholders and analysis code. Specifically it:
1) Opens a docx (or exports one from a Google Doc)
2) Searches for placeholder tags containted in curly brackets {}.
3) For each placeholder tag, the code identifies whether an object (e.g. string or number) exists in the R environment with the same name.
4) If it does, it replaces the text of the placeholder tag with the value of the object in the R environment.

Figures are similarly designated with curly brackets in the text, but in code are designated with the bundle_ggplot() function to wrap together both the ggplot underlying the figure and dimensional settings.
