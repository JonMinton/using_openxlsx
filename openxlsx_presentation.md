Why and how to use openxlsx
========================================================
author: Jon Minton
date: 21 Feb 2020
autosize: true

Introduction: The problem
========================================================

- Markdown has `kable` etc
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
    - Sharepoint (aaagh!)
    - Security concerns?
    - Could then save/export to MS (but manual process)
    - And I've not really used it!
    
***

- [`xlsx`](https://cran.r-project.org/web/packages/xlsx/xlsx.pdf): Java-dependent
    - Java increasingly not supported
    - IT permissions required
- [`openxslx`](https://cran.r-project.org/web/packages/openxlsx/index.html): *Not* Java dependent
    - Read
    - write
    - formatting/editing
    
Basic workflow: Using `openxlsx` quickly
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

Basic workflow: Using `openxlsx` properly
=========================================================

Minimum of four stages 

1. `createworkbook()`
2. `addWorksheet()`
3. `WriteData()` or `writeDatatable()`
4. `saveWorkbook()`


1. `createWorkbook()`
==============================

1. `createWorkbook()`: Create skeleton workbook
    - Just within Rstudio session
    - Not saved as file
    - Must be passed to object to be worked on later


```r
wb <- createWorkbook()
```
    
2. `addWorksheet()`
=======================================

2. `addWorksheet(wb, [other arguments])`: Attaches a worksheet to the `wb` object
    - Still not saved!
    - Still no contents yet!
    - `addWorksheet(wb, sheetName = "my sheet name")`


```r
wb <- createWorkbook()
addWorksheet(wb, sheetName = "my sheet name")
```

3. `writeData()`/`writeDatatable()`
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


4. `saveWorkbook()`
==================================


```r
wb <- createWorkbook()
addWorksheet(wb, sheetName = "my sheet name")
writeData(wb, sheet = 1, x = mtcars)
# writeData(wb, sheet = "my sheet name", x = mtcars)
saveWorkbook(wb, "basics.xlsx")
```

4. `saveWorkbook()`
====================================

- Add `overwrite=`




```
Error in (function (filename = "Rplot%03d.png", width = 480, height = 480,  : 
  unable to start png() device
```
