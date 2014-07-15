
# -------------------------------------------------------------
#
#   ReporteRs version: Compare methods of creating docx files
#
#   Richard Layton
#   July 2014
#
#   requires packages: reshape2, stringr, plyr, 
#       RColorBrewer, lattice, ReporteRs
#
# -------------------------------------------------------------

# docx---docx---docx---docx---docx---docx---docx---docx---docx

### SET UP

library(ReporteRs)
options(
  "ReporteRs-fontsize" = 11
  , "ReporteRs-default-font" = "Palatino Linotype"
)

# name the docx file. 
# I'm using relative directories inside an RStudio project
docx.file <- "Reports/OutputDocx_ReporteRs.docx"

# assign the template
doc <- docx(template = "Templates/Template_ReporteRs.docx")

# display available styles
styles(doc)

# date
doc <- addParagraph(doc, "July 13, 2014", stylename = "Normal")

# addressee
doc <- addParagraph(doc, "To:   My R friends", stylename = "NoSpacing")
doc <- addParagraph(doc
                   , "Re:   Comparing methods of creating docx files"
                   , stylename = "Normal")

# introduction
doc <- addParagraph(doc, "The usual report I create for my collaborators includes data tables, graphs, and short discussions. In this test documant, I've included some of each, formatted as closely as possble to my usual standards for document design.", stylename = "Normal")

# heading 1
doc <- addParagraph(doc, "Creating a docx file using ReporteRs", stylename = "Heading1")

doc <- addParagraph(doc, "I first created an R file to manipulate data and create a graph. With the analysis complete, I added in the ReporteRs functions.", stylename = "Normal")

doc <- addParagraph(doc, "In a separate step, I set up a docx template file, changed its styles, and added a logo in the header of the first page. ", stylename = "Normal")

# bulleted list intro
doc <- addParagraph(doc, "Some thoughts:", stylename = "Normal")
# list items
items <- c(
  "No spellcheck for the R file. And for an unknown reason, Word spell check is not finding misspelled words in the docx created by ReporteRs."
  , "Many mouse clicks could be eliminated if the docx file could be overwritten while open. Or perhaps functions could be added to close an open file, overwrite it, and reopen it."
  , "I could not figure out how to easily create an em-dash."
)
# make the list
doc <- addParagraph(doc, value = items, stylename = "BulletList")



# Rcode---Rcode---Rcode---Rcode---Rcode---Rcode---Rcode---Rcode

###  DATA

# the 1940 Virginia deaths data, converted to a DF
rate <- c(VADeaths)
n <- length(rate)
df1 <- data.frame(
  rate = rate
  , age = rep(ordered(rownames(VADeaths)), length.out = n)
  , sex = gl(2, 5, n, labels = c("Male", "Female"))
  , site =  gl(2, 10, n, labels = c("Rural", "Urban"))
)

# manipulate the DF to wide form for printout later
df2 <- df1
df2$rate <- round(df2$rate, 0)

# cast to wide form for table
library(reshape2)
df2 <- dcast(df2, age ~ site + sex, value.var = "rate")

# edit col headings
library(plyr)
df2 <- rename(
  df2
  , c("age" = "Age group"
      , "Rural_Male" = "Rural male"
      , "Rural_Female" = "Rural female"
      , "Urban_Male" = "Urban male"
      , "Urban_Female" = "Urban female"
  )
)

# df1 is used for graph
# df2 is used for printout

### GRAPH

# obtain color palettes
library(RColorBrewer)
# divergent Brown-BlueGreen
BrBG    <- brewer.pal(5,"BrBG")
darkBr  <- BrBG[1]
lightBr <- BrBG[2]
lightBG <- BrBG[4] 
darkBG  <- BrBG[5]
# grays
Grays <- brewer.pal(6,"Greys")
gray1 <- Grays[1]
gray2 <- Grays[2]
gray3 <- Grays[3]
gray4 <- Grays[4]
gray5 <- Grays[5] 
gray6 <- Grays[6]

# assign two colors for distinguishing Rural from Urban
col.array <- c(darkBG, darkBr) 

# initialize graph settings
library(lattice)
myGroups <- df1$site
selLabel <- c(2, 7) # for adding text labels 
xTicksAt <- seq(10, 70, 10)
yTicksAt <- 1:length(levels(df1$age))

# to set all typefaces the same
allFontSet <- list(font = 1, cex = 1, fontfamily = 'sans') 

# graph settings
myTheme <- list(
  fontsize = list(text = 8, points = 5)
  , add.text = allFontSet # strip label
  , axis.line = list(col = gray4, lwd = 0.5)
  , axis.text = allFontSet
  , strip.border = list(col = gray4, lwd = 0.5) 
  , layout.heights = list(strip = 1.25)
  , strip.background = list(col = gray2 )  	
)
myScales <- list(
  x = list(at = xTicksAt)
  , y = list(alternating = FALSE)
  , tck = c(0.8, 0.8)
)
myPanel <- function(x, y, ...){
  panel.abline( # grid lines
    h = yTicksAt, v = xTicksAt, ... 
    , col = gray2
    , lwd = 0.5
  )
  panel.superpose( # data markers and lines
    x, y, ... 
    , pch = 21
    , cex = 1.1
    , col = col.array
    , fill = col.array
    , type = "o"
  )
  panel.text( # labels
    x = x[selLabel], y = y[selLabel], ... 
    , pos = c(2, 4)
    , offset = 1
    , labels = myGroups[selLabel+c(0,10)]
    , col = col.array
  )
}

# create the graph
f1 <- xyplot(
  age ~ rate | reorder(sex, rate, median)
  , data = df1
  , xlab = "Death rate (per 1000)"
  , ylab = list(label = "Age group", rot = 0)
  , panel = myPanel
  , layout = c(1, NA)
  , groups = myGroups
  , scales = myScales
  , par.settings = myTheme
)

# print figure to RStudio pane
print(f1)



# docx---docx---docx---docx---docx---docx---docx---docx---docx

### DATA TABLE 

# section heading
doc <- addTitle(doc, "Data source", level = 1)

# paragraph with a word italicized
pot1 <- "The" + pot(" VADeaths ", textProperties(font.style = "italic")) + "data are furnished in the base R install."
doc <- addParagraph(doc, set_of_paragraphs(pot1), stylename="Normal" )

doc <- addParagraph(doc, "Learning to manage the table formatting for the first time was as about as involved as learning it in LaTeX for the first time---difficult, but once a design is known, the commands can be reused in the next document.", stylename = "Normal")

# Add a caption above the table
doc <- addPageBreak(doc) # manual pagination
doc <- addParagraph(doc, value = "Death rate data (per 1000), Virginia 1940", stylename = "rTableLegend")

# Create a FlexTable 
MyFTable <- FlexTable(
  data = df2
  , add.rownames = FALSE
  , body.text.props = textProperties(font.weight = "normal"
                                     , font.size = 10
                                     , font.family = "Calibri")
  , header.text.props = textProperties(font.weight = "normal"
                                       , font.size = 10
                                       , font.family = "Calibri")
  , body.par.props = parProperties(text.align = "right")
  , header.par.props = parProperties(text.align = "right")
)

# applies a border grid on table
MyFTable <- setFlexTableBorders(
  MyFTable
  , inner.vertical = borderProperties(style = "none")
  , inner.horizontal = borderProperties(style = "none")
  , outer.vertical = borderProperties(style = "none")
  , outer.horizontal = borderProperties(color = gray5, style = "solid")
)

# add MyFTable into document 
doc <- addFlexTable(doc, MyFTable)



### GRAPH

# section heading
doc <- addTitle(doc, "Data display", level = 1)

# paragraph with a word italicized
pot1 <- "The figure is drawn using the lattice package. The size of the figure is controlled using ResporteRs" + pot(" addPlot() ", textProperties(font.style = "italic")) + "function."

doc <- addParagraph(doc, set_of_paragraphs(pot1), stylename="Normal" )





# add the graph
doc <- addPlot(doc = doc, fun = print, x = f1 
              , width = 5.3
              , height = 4
)

# caption
doc <- addParagraph(doc, value = "Comparing male and female death rates in rural and urban Virginia in 1940.", stylename = "rPlotLegend")

# add an extra space between the caption and the next paragraph 
doc <- addParagraph(doc, " ", stylename = "NoSpacing" )

# discuss 
doc <- addParagraph(doc, "Rates are nearly identical for rural and urban females, with a systematic increase among rural males and a further increase for urban males.", stylename = "Normal" )





### CREATE THE DOCX FILE

writeDoc(doc, file = docx.file)

# last line
