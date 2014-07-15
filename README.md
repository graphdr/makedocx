makedocx
========

Comparing two methods of using R to create docx files for collaborators.

Like many R users, I collaborate with colleagues who prefer docx and pptx file formats. My usual work flow has been to use RStudio/knitr/LaTeX (.Rnw files) to produce a PDF document that I send ’round to my colleagues for review and comment. While *my* work is reproducible, PDFs are not the most convenient format for iterative review and comment.

At the user!2014 conference, I learned of two methods to create docx files dynamically: rmarkdown and ReporteRs. I worked with both packages for a couple of days and can offer this initial assessment. In brief, ReporteRs provides superior control of document design but requires additional (and somewhat awkward, though effective) markup; rmarkdown provides less control of document design but is more human-readable, consistent with literate programming principles.

The files are organized in three directories:
*  Reports. Used for the source code and the output docxfiles.
*  Visuals. For graphics files.
*  Templates. For the docx template files used by ReporteRs and RMarkdown. 

On my laptop, these files are located in a directory assigned to an RStudio Project. The user will have to edit the relative paths in the the two files, *SourceFile_ReporteRsToDocx.R* and *SourceFile_RMarkdownToDocx.Rmd*. 



