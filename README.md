mschart R package
================

<!-- README.md is generated from README.Rmd. Please edit that file -->
[![Project Status: WIP - Initial development is in progress, but there has not yet been a stable, usable release suitable for the public.](http://www.repostatus.org/badges/latest/wip.svg)](http://www.repostatus.org/#wip) [![Build Status](https://travis-ci.org/ardata-fr/mschart.svg?branch=master)](https://travis-ci.org/ardata-fr/mschart) [![AppVeyor Build Status](https://ci.appveyor.com/api/projects/status/github/ardata-fr/mschart?branch=master&svg=true)](https://ci.appveyor.com/project/ardata-fr/mschart)

The `mschart` package provides a framework for easily create charts for 'Microsoft PowerPoint' documents.

Installation
------------

You can install the package from github with:

``` r
# install.packages("devtools")
devtools::install_github("ardata-fr/mschart")
```

Example
-------

This is a basic example which shows you how to create a line chart.

``` r
library(mschart)
library(officer)

linec <- ms_linechart(data = iris, x = "Sepal.Length",
                      y = "Sepal.Width", group = "Species")
linec <- chart_ax_y(linec, num_fmt = "0.00", rotation = -90)
```

Then use package `officer` to send the object as a chart.

``` r
doc <- read_pptx()
doc <- add_slide(doc, layout = "Title and Content", master = "Office Theme")
doc <- ph_with_chart(doc, value = linec)

print(doc, target = "example.pptx")
```

Details
-------

The following objects are available:

-   barcharts with function `ms_barchart()`
-   linecharts with function `ms_linechart()`
-   scatter plots with function `ms_scatterchart()`
-   areacharts with function `ms_areachart()`

All these functions are returning *chart* objects that can be manipulated.

First, you should use method `chart_settings()`. Parameters are specific to each type of chart.

``` r
my_barchart <- ms_barchart(data = browser_data, 
  x = "browser", y = "value", group = "serie")

my_barchart <- chart_settings( my_barchart, 
  dir="vertical", grouping="stacked",
  gap_width = 150, overlap = 100 )
```

You can then customise axes with functions `chart_ax_x`, `chart_ax_y`:

``` r
my_barchart <- chart_ax_x(my_barchart, cross_between = 'between', 
  major_tick_mark = "in", minor_tick_mark = "none")
my_barchart <- chart_ax_y(my_barchart, num_fmt = "0.00", rotation = -90)
```

To add titles, use function `chart_labels`:

``` r
my_barchart <- chart_labels(my_barchart, title = "A main title", 
  xlab = "x axis title", ylab = "y axis title")
```

To modify fill, stroke colours, symbols and size of symbols associated with series, use the following functions: `chart_data_fill`, `chart_data_stroke`, `chart_data_size`, `chart_data_symbol`.

``` r
my_scatter <- chart_data_fill(my_barchart,
  values = c(serie1 = "#6FA2FF", serie2 = "#FF6161", serie3 = "#81FF5B") )
```

Inspired from `ggplot2`, there is a `set_theme` function. It let customise grid lines, titles formatting properties, etc.

``` r
mytheme <- mschart_theme(
  axis_title_x = fp_text(color = "red", font.size = 24, bold = TRUE),
  axis_title_y = fp_text(color = "green", font.size = 12, italic = TRUE),
  grid_major_line_y = fp_border(width = 1, color = "orange"),
  axis_ticks_y = fp_border(width = 1, color = "orange") )
my_barchart <- set_theme(my_barchart, mytheme)
```