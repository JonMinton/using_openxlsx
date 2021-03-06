Why and how to use openxlsx
========================================================
author: Jon Minton
date: 21 Feb 2020
autosize: true

Introduction: The problem
========================================================

- Markdown has `kable`/`kableExtra` etc
    - Renders nicely as HTML
    - Doesn't copy over to Excel /Word too easily
    - Manual and error prone process
- Legacy demand for tablular outputs in Excel/Word
    - Reports
    - Scientific articles (except in harder science fields)
- How to produce these in a way that meets modern automation/reproducibility standards and aspirations?

Some (partial) solutions
==========================================================



- [`readxl`](https://readxl.tidyverse.org/): Reading only 
- [`googlesheets`](https://cran.r-project.org/web/packages/googlesheets/vignettes/basic-usage.html): We're more wedded to M$ than Google
    - Sharepoint :(
    - Security concerns?
    - Could then save/export to MS (but manual process)
    - And I've not really used it!
    
***

- [`xlsx`](https://cran.r-project.org/web/packages/xlsx/xlsx.pdf): Java-dependent
    - Java increasingly not supported
    - IT permissions required
- [`openxlsx`](https://cran.r-project.org/web/packages/openxlsx/index.html): *Not* Java dependent
    - Read
    - write
    - formatting/editing
    
Basic workflow: Using openxlsx quickly
=========================================================

From [Online Introduction](https://cran.r-project.org/web/packages/openxlsx/vignettes/Introduction.pdf)


```r
# Workbook with given name with worksheet containing iris
write.xlsx(iris, file = "writeXLSX1.xlsx") 
write.xlsx(iris, file = "writeXLSXTable1.xlsx", asTable = TRUE) 
# As excel table (slightly different formatting/referencing defaults)
```


```r
## write a list of data.frames to individual worksheets using list names as worksheet names
l <- list("IRIS" = iris, "MTCARS" = mtcars)
write.xlsx(l, file = "writeXLSX2.xlsx")
```

Basic workflow: Using openxlsx properly
=========================================================

Minimum of four stages 

1. `createworkbook()`
2. `addWorksheet()`
3. `WriteData()` or `writeDatatable()`
4. `saveWorkbook()`


1. createWorkbook()
==============================

1. `createWorkbook()`: Create skeleton workbook
    - Just within Rstudio session
    - Not saved as file
    - Must be passed to object to be worked on later


```r
wb <- createWorkbook()
```
    
2. addWorksheet()
=======================================

2. `addWorksheet(wb, [other arguments])`: Attaches a worksheet to the `wb` object
    - Still not saved!
    - Still no contents yet!
    - `addWorksheet(wb, sheetName = "my sheet name")`


```r
wb <- createWorkbook()
addWorksheet(wb, sheetName = "my sheet name")
```

3. writeData() / writeDatatable()
================================

3. `writeData()` or `writeDatatable()`: Add data to worksheet in workbook.
    - *Still* not saved!
    - Now need to specify both `wb` object name and sheet name within `wb`
    

```r
wb <- createWorkbook()
addWorksheet(wb, sheetName = "my sheet name")
writeData(wb, sheet = 1, x = mtcars)
# writeData(wb, sheet = "my sheet name", x = mtcars)
```

- Additional arguments possible
    - Formatting: `borders=`, etc
    - Positioning: `startCol=`, `startRow=`, etc


4. saveWorkbook()
==================================


```r
wb <- createWorkbook()
addWorksheet(wb, sheetName = "my sheet name")
writeData(wb, sheet = 1, x = mtcars)
# writeData(wb, sheet = "my sheet name", x = mtcars)
saveWorkbook(wb, "basics.xlsx")
```

4. saveWorkbook()
====================================

- Add `overwrite=`


```r
wb <- createWorkbook()
addWorksheet(wb, sheetName = "my sheet name")
writeData(wb, sheet = 1, x = mtcars)
# writeData(wb, sheet = "my sheet name", x = mtcars)
saveWorkbook(wb, "basics.xlsx", overwrite = TRUE)
```


Between (1) and (4)
====================================

Many options for further customisation of worksheet contents before saving 

- Merging columns and rows
- Adding cell colours
- Changing fonts
- Resizing column widths and row heights
- Conditional formatting 
- etc 
- etc 
- etc

Examples used in practice
====================================

1. Simple tables using ONS/HMD data 
    - Cell merging
    - Number of DPs displayed 
    - Resizing
    - Font formatting
    - Border lines
2. Heatmaps of correlations (HMD)
    - Multiple worksheets
    - Conditional formatting: values -> colour
    - Resized to fit on one page 

Example 1: Simple tables 
====================================

- An example in github repo [`Bayes_Factor_Slowdown`](https://github.com/JonMinton/bayes_factor_slowdown)
    - Specific example [here](https://github.com/JonMinton/bayes_factor_slowdown/blob/master/markdown/bayes_factor_ons_e0.Rmd#L243-L317)
    - Produces [this file](https://github.com/JonMinton/bayes_factor_slowdown/blob/master/bayes_paper/tables/M01_hmd_countries_decade.xlsx)

Example 1: Reflections
====================================

- [This function](https://github.com/JonMinton/bayes_factor_slowdown/blob/master/markdown/bayes_factor_ons_e0.Rmd#L248-L252) is a bit of a hack
- Could put all tables in a single workbook
    - But more likely to forget to 'close the door' after opening it
- Formatting code quite extensive
    - But this reflects genuine complexity in number of specific processes and instructions involved in Excel formatting
    
Example 2: Heatmaps 
====================================

- From [this github repo](https://github.com/JonMinton/rising_tide)
    - Specific example [within here](https://github.com/JonMinton/rising_tide/blob/master/analyses.Rmd)
    - What it shows
        - Correlation between trends in log mortality at two different individual ages (?!?!)
        - **'Tartanplots!'**
    - Motivation: to produce an Excel version of [this kind of heatmap](https://raw.githubusercontent.com/JonMinton/rising_tide/master/figures/heatmap_2011_onwards.png)
        - To look up values (for commentary)
        - To slice through different time periods 

***

- Version with just a couple of slices
    - The [Code is here](https://github.com/JonMinton/rising_tide/blob/master/analyses.Rmd#L1221-L1305)
    - Produces [this worksheet](https://github.com/JonMinton/rising_tide/blob/master/tables/blocks_both_main_periods.xlsx)
- Version with many slices (25 year period, 5 year rolling starts)
    - [Code here](https://github.com/JonMinton/rising_tide/blob/master/analyses.Rmd#L1310-L1403)
    - Produces [this behemoth](https://github.com/JonMinton/rising_tide/blob/master/tables/many_blocks.xlsx)!


Example 2: Reflections 
====================================

- Benefits of automating now stars to pay more clearly for cost of thinking/expressing formatting decisions formally
- Can get lost in/swamped by the data itself!
- Proof of principle that 10s/100s/1000s of worksheets can be produced at once


Summary
=====================================

- Strengths
    - Fully automated
         - Fewer opportunities to introduce errors
         - Change the data -> change the excel output
    - Forces clear thinking about formatting/styling
    - No Java dependencies
    - Simple save options also available
- Limitations/development opportunities
    - xlsx expressions focus too much attention on the programmatic cf conceptual
        - Dirty hacks (e.g. when merging multiple groups of cells)
        - kable -> Excel workbook table converter
        
        
Suggestions 
===================================

- Could add more metadata/descriptive data procedurally too
    - Options to add hyperlinks etc? 
- **Cost-benefit question**: Cost of thinking < benefits of scaling?
- Some alternatives exist
     - Google Sheets? 
     - Convenience function library to develop? 
- Preference for Excel workbooks not going any time soon
- This seems worth investing in to meet people with these preferences


