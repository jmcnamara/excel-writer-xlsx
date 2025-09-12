# Changelog

This is a changelog for Excel::Writer::XLSX.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).


## [1.15] - 2025-11-12

### Added

- Add the option to position custom data labels in the same way that the data
  labels can be positioned for the entire series.

- Add border, fill, gradient and pattern formatting options for
  chart titles and also chart axis titles.


### Fixed

- Fix failed match for range string in conditional format.
  Strings used in a cell equality must be quoted. In order to
  ensure this there is a check for non-range/non-numeric strings.
  However this test wasn't anchored properly and gave a false
  positive.

  Issue [#312](https://github.com/jmcnamara/excel-writer-xlsx/issues/312).

- Fixed an issue where `set_row()` with the default height of 15 was
  ignored if `set_default_row()` was also used.

- Fixed issue with 0 DPI in PNG images.


## [1.14] - 2024-10-21

### Fixed

- Cleaned up release tarball to remove editor dot files.


## [1.13] - 2024-10-13

### Fixed

- Fixed issue with html color for border colors.

  Issue [#302](https://github.com/jmcnamara/excel-writer-xlsx/issues/302) and
  [#305](https://github.com/jmcnamara/excel-writer-xlsx/issues/305).


## [1.12] - 2024-02-26

### Added

- Added support for embedding images into worksheets with
  worksheet `embed_image()`.
  This can be useful if you are building up a spreadsheet of products with
  a column of images for each product. Embedded images move with the cell
  so they can be used in worksheet tables or data ranges that will be
  sorted or filtered.
  This functionality is the equivalent of Excel's menu option to insert an
  image using the option to "Place in Cell" which is available in Excel
  365 versions from 2023 onwards.

- Added support for Excel 365 `IMAGE()` future.

- Added trendline equation formatting for Charts.

- Added support for leader lines to all chart types.

- Added chart option to display `N/A` as empty cells.

- Add support for `invert_if_negative` color option in Charts.

- Added worksheet `very_hidden()` method to hide a worksheet in a way that
  it can only be unhidden by VBA. Feature Request [#228](https://github.com/jmcnamara/excel-writer-xlsx/issues/228).


### Fixed

- Fixed indentation and alignment property mismatch.
  Fix issue where a horizontal alignment format was ignored if the
  indentation was also set.


## [1.11] - 2023-03-19

### Added

- Added support for simulated worksheet `autofit()`.

- Refactored internal column property handling to allow column ranges
  to be overridden (a common UX expectation).

- Add `quote_prefix` format property.


### Fixed

- Add fix for forward/backward slash issue on Windows perl builds.
  Possibly introduced by a change in File::Find. Issue [#284](https://github.com/jmcnamara/excel-writer-xlsx/issues/284)

- Fix for duplicate number formats. Issue [#283](https://github.com/jmcnamara/excel-writer-xlsx/issues/283)

- Add fix for worksheets with tables and background images.

- Replace/fix the worksheet protection password algorithm
  so that is works correctly for strings over 24 chars.


## [1.10] - 2022-12-30

### Added

- Add support for new Excel 365 dynamic functions.


## [1.09] - 2021-05-14

### Added

- Added support for background images in worksheets. See set_background().

- Added support for GIF image files (and in Excel 365, animated GIF files).

- Added support for pixel sizing in set_row() and set_column() via new
  functions called set_row_pixels() and set_column_pixels().

- Added initial support for dynamic arrays in formulas.


## [1.08] - 2021-03-31

### Added

- Added ability to add accessibility options "description" and
  "decorative" to images via insert_image().

- Added the workbook read_only_recommended() method to set the Excel
  "Read-only Recommended" option that is available when saving a file.

- Added option to set a chart crossing to 'min' as well as the existing
  'max' option. The 'min' option isn't available in the Excel interface
  but can be enabled via VBA.

- Added option to unprotect ranges in protected worksheets.

- Added check, and warning, for worksheet tables with no data row.  Either
  with or without a header row.

- Added ignore_errors() worksheet method to ignore Excel worksheet
  errors/warnings in user defined ranges.


### Fixed

- Fixed issue where pattern formats without colors where given a default
  black fill color.

- Fix issue where custom chart data labels didn't inherit the position for
  the data labels in the series.

- Fixed issue with relative url links in images.

- Fixed issue where headers/footers were restricted to 254 characters
  instead of 255.


## [1.07] - 2020-08-06

### Added

- Added support for Border, Fill, Pattern and Gradient formatting to chart
  data labels and chart custom data labels.


## [1.06] - 2020-08-03

### Fixed

- Fix for issue where array formulas weren't included in the output file
  for certain ranges/conditions.


## [1.05] - 2020-07-30

### Added

- Added support for custom data labels in charts.


## [1.04] - 2020-05-31

### Added

- Added support for "stacked" and "percent_stacked" Line charts.


### Fixed

- Fix for worksheet objects (charts and images) that are inserted with an
  offset that starts in a hidden cell.

- Fix for issue with default worksheet VBA codenames.


### Removed

- Removed error in add_worksheet() for sheet name "History" which is a
  reserved name in English version of Excel. However, this is an allowed
  worksheet name in some Excel variants so the warning has been turned into
  a documentation note instead.


## [1.03] - 2019-12-26

### Added

- Fix for duplicate images being copied to an Excel::Writer::XLSX
  file. Excel uses an optimization where it only stores one copy of a
  repeated/duplicate image in a workbook. Excel::Writer::XLSX didn't do
  this which meant that the file size would increase when then was a large
  number of repeated images. This release fixes that issue and replicates
  Excel's behavior.


## [1.02] - 2019-11-07

### Added

- Added support for hyperlinks in worksheet images.

- Increased allowable url length from 255 to 2079 characters, as allowed in
  more recent versions of Excel.


## [1.01] - 2019-10-28

### Added

- Added support for stacked and  East Asian vertical chart fonts.

- Added option to control positioning of charts or images when cells are
  resized.

- Added support for combining Pie/Doughnut charts.


### Fixed

- Fixed sizing of cell comment boxes when they cross columns/rows that have
  size changes that occur after the comment is written. Comments should
  now behave like other worksheet objects such as images and charts.

- Fix for structured reference in chart ranges.


## [1.00] - 2019-04-07

### Fixed

- Fixed issue where images that started in hidden rows/columns weren't placed
  correctly in the worksheet.

- Fixed the mime-type reported by system "file(1)". The mime-type reported
  by "file --mime-type"/magic was incorrect for Excel::Writer::XLSX files
  since it expected the "[Content_types]" to be the first file in the zip
  container.


## [0.99] - 2019-02-10

### Added

- Added font and font_size parameters to write_comment().

- Allow formulas in date field of data_validation().

- Added top_left chart legend position.

- Added legend formatting options.

- Added set_tab_ratio() method to set the ratio between the worksheet tabs
  and the horizontal slider.

- Added worksheet hide_row_col_headers() method to turn off worksheet row
  and column headings.

- Add functionality to align chart category axis labels.


### Fixed

- Fix for issue with special characters in worksheet table functions.

- Fix handling of 'num_format': '0' in duplicate formats.


## [0.98] - 2018-04-14

### Added

- Set the xlsx internal file member datetimes to 1980-01-01 00:00:00 like
  Excel so that apps can produce a consistent binary file once the
  workbook set_properties() created date is set. Feature request [#168](https://github.com/jmcnamara/excel-writer-xlsx/issues/168).


## [0.97] - 2018-04-10

### Added

- Added Excel 2010 data bar features such as solid fills and control over
  the display of negative values.

- Added default formatting for hyperlinks if none is specified. The format
  is the Excel hyperlink style so links change color after they are
  clicked.


### Fixed

- Fixed missing plotarea formatting in pie/doughnut charts.


## [0.96] - 2017-09-16

### Added

- Added icon sets to conditional formatting. Feature request [#116](https://github.com/jmcnamara/excel-writer-xlsx/issues/116).


## [0.95] - 2016-06-13

### Added

- Added workbook set_size() method. Feature request [#59](https://github.com/jmcnamara/excel-writer-xlsx/issues/59).


## [0.94] - 2016-06-07

### Added

- Added font support to chart tables.

  Feature request [#96](https://github.com/jmcnamara/excel-writer-xlsx/issues/96).


## [0.93] - 2016-06-07

### Added

- Added trendline properties: intercept, display_equation and
  display_r_squared. Feature request [#153](https://github.com/jmcnamara/excel-writer-xlsx/issues/153).


## [0.92] - 2016-06-01

### Fixed

- Fix for insert_image issue when handling images with zero dpi.


## [0.91] - 2016-05-31

### Added

- Add set_custom_property() workbook method to set custom document
  properties.


## [0.90] - 2016-05-13

### Added

- Added get_worksheet_by_name() workbook method to retrieve a worksheet
  in a workbook by name.

  Feature request [#124](https://github.com/jmcnamara/excel-writer-xlsx/issues/124).


### Fixed

- Fixed issue where internal file creation and modification dates where
  in the local timezone instead of UTC.

  Issue [#162](https://github.com/jmcnamara/excel-writer-xlsx/issues/162).

- Fixed issue with "external:" urls with space in sheetname.

- Fixed issue where Unicode full-width number strings were treated as
  numbers in write().

  Issue [#160](https://github.com/jmcnamara/excel-writer-xlsx/issues/160).


## [0.89] - 2016-04-16

### Added

- Added write_boolean() worksheet method to write Excel boolean values.


## [0.88] - 2016-01-14

### Added

- Added transparency option to solid fills in chart areas.

- Added options to configure chart axis tick placement.

  Feature request [#158](https://github.com/jmcnamara/excel-writer-xlsx/issues/158).


## [0.87] - 2016-01-12

### Added

- Added chart pattern and gradient fills.

- Added option to set chart tick interval.

- Add checks for valid and non-duplicate worksheet table names.

- Added support for table header formatting and a fix for wrapped lines in
  the header.


## [0.86] - 2015-10-19

### Fixed

- Fix to allow chartsheets to support combined charts.

- Fix for images with negative offsets.


### Added

- Allow hyperlinks longer than 255 characters when the link and anchor
  are each less than or equal to 255 characters.

- Added hyperlink_base document property.

- Added option to allow data validation input messages with the ‘any’
  validate parameter.

  Issue [#144](https://github.com/jmcnamara/excel-writer-xlsx/issues/144).

- Added "stop if true" feature to conditional formatting.

  Issue [#138](https://github.com/jmcnamara/excel-writer-xlsx/issues/138).

- Added better support and documentation for html colors throughout
  the module. The use of the Excel97 color palette is supported for
  backward compatibility but deprecated.

  Issue [#97](https://github.com/jmcnamara/excel-writer-xlsx/issues/97).


## [0.85] - 2015-08-05

### Fixed

- Fixes for new redundant sprintf arguments warnings in perl 5.22.

  Issue [#134](https://github.com/jmcnamara/excel-writer-xlsx/issues/134).

- Fix url encoding of links to external files and dirs.


## [0.84] - 2015-04-21

### Added

- Added support for chart axis display units (thousands, million, etc.).

- Added option to set printing in black and white. Issue [#125](https://github.com/jmcnamara/excel-writer-xlsx/issues/125).

- Added chart styles example.

- Added gradient fill support.

- Added support for clustered charts.

- Added support for boolean error codes.
  0.83 2015-3-17

- Added option to combine two different chart types. For example to
  create a Pareto chart.
  0.82 2015-3-14

- Added extra documentation on how to handle VBA macros and added
  automatic and manual setting of workbook and worksheet VBA codenames.

  Issue [#60](https://github.com/jmcnamara/excel-writer-xlsx/issues/60).

- Fix for set_start_page() for values > 1.

- Fix to copy user defined chart properties, such as trendlines,
  so that they aren't overwritten.

  Issue [#121](https://github.com/jmcnamara/excel-writer-xlsx/issues/121).

- Added column function_value option to add_table to allow
  function value to be set.

- Allow explicit text categories in charts.

  Issue [#102](https://github.com/jmcnamara/excel-writer-xlsx/issues/102)

- Fix for column/bar gap/overlap on y2 axis.

  Issue [#113](https://github.com/jmcnamara/excel-writer-xlsx/issues/113).


## [0.81] - 2014-11-01

### Added

- Added chart axis line and fill properties.


## [0.80] - 2014-10-29

### Added

- Chart Data Label enhancements. Added number formatting, font handling
  (issue #106), separator (issue #107) and legend key.

- Added chart specific handling of data label positions since not all
  positions are available for all chart types. Issue [#110](https://github.com/jmcnamara/excel-writer-xlsx/issues/110).


## [0.79] - 2014-10-16

### Added

- Added option to add images to headers and footers.

- Added option to not scale header/footer with page.


### Fixed

- Fixed issue where non 96dpi images were not scaled properly in Excel.

- Fix for issue where X axis title formula was overwritten by the
  Y axis title.


## [0.78] - 2014-09-28

### Added

- Added Doughnut chart with set_rotation() and set_hole_size()
  methods.

- Added set_rotation() method to Pie charts.

- Added set_calc_mode() method to control automatic calculation of
  formulas when worksheet is opened.


## [0.77] - 2014-05-06

### Fixed

- Fix for incorrect chart offsets in insert_chart() and set_size().
  Reported by Kevin Gilpin.


## [0.76] - 2013-12-31

### Added

- Added date axis handling to charts.

- Added support for non-contiguous chart ranges.


### Fixed

- Fix to remove duplicate set_column() entries.


## [0.75] - 2013-12-02

### Added

- Added interval unit option for category axes.


### Fixed

- Fix for axis name font rotation. Issue [#83](https://github.com/jmcnamara/excel-writer-xlsx/issues/83).

- Fix for several minor issues with Pie chart legends.


## [0.74] - 2013-11-17

### Fixed

- Improved defined name validation.

  Issue [#82](https://github.com/jmcnamara/excel-writer-xlsx/issues/82).


### Added

- Added set_title() option to turn off automatic title.

  Issue [#81](https://github.com/jmcnamara/excel-writer-xlsx/issues/81).

- Allow positioning of plotarea, legend, title and axis names.

  Issue [#80](https://github.com/jmcnamara/excel-writer-xlsx/issues/80).


### Fixed

- Fix for modification of user params in conditional_formatting().

  Issue [#79](https://github.com/jmcnamara/excel-writer-xlsx/issues/79).

- Fix for star style markers.


## [0.73] - 2013-11-08

### Added

- Added custom error bar option to charts.


### Fixed

- Fix for tables added in non-sequential order.

- Fix for scatter charts with markers on non-marker series.


## [0.72] - 2013-08-28

### Fixed

- Fix for charts and images that cross rows and columns that are
  hidden or formatted but which don’t have size changes.


## [0.71] - 2013-08-24

### Fixed

- Fixed issue in image handling.

- Added fix to ensure formula calculation on load regardless of
  Excel version.


## [0.70] - 2013-07-30

### Fixed

- Fix for rendering images that are the same size as cell boundaries.
  GitHub issue #70.

- Added fix for inaccurate column width calculation.


### Added

- Added Chart line smoothing option.


## [0.69] - 2013-06-12

### Added

- Added chart font rotation property. Mainly for use with date axes
  to make the display more compact.


### Fixed

- Fix for 0 data in Worksheet Tables. [#65](https://github.com/jmcnamara/excel-writer-xlsx/issues/65).
  Reported by David Gang.


## [0.68] - 2013-06-06

### Fixed

- Fix for issue where shapes on one worksheet corrupted charts on a
  subsequent worksheet. [#52](https://github.com/jmcnamara/excel-writer-xlsx/issues/52).

- Fix for issue where add_button() invalidated cell comments in the
  same workbook. [#64](https://github.com/jmcnamara/excel-writer-xlsx/issues/64).


## [0.67] - 2013-05-06

### Fixed

- Fix for set_selection() with cell range.


## [0.66] - 2013-04-12

### Fixed

- Fix for issue with image scaling.


## [0.65] - 2012-12-31

### Added

- Added options to format series Gap/Overlap for Bar/Column charts.


## [0.64] - 2012-12-22

### Added

- Added the option to format individual points in a chart series.
  This allows Pie chart segments to be formatted.


## [0.63] - 2012-12-19

### Added

- Added Chart data tools such as:
  Error Bars
  Up-Down Bars
  High-Low Lines
  Drop Lines.
  See the chart_data_tool.pl example.


## [0.62] - 2012-12-12

### Added

- Added option for adding a data table to a Chart X-axis.
  See output from chart_data_table.pl example.


## [0.61] - 2012-12-11

### Added

- Allow a cell url string to be over written with a number or formula
  using a second write() call to the same cell. The url remains intact.

  Issue [#48](https://github.com/jmcnamara/excel-writer-xlsx/issues/48).

- Added set_default_row() method to set worksheet default values for
  rows.

- Added Chart set_size() method to set the chart dimensions.


## [0.60] - 2012-12-05

### Added

- Added Excel form buttons via the worksheet insert_button() method.
  This allows the user to tie the button to an embedded macro imported
  using add_vba_project().
  The portal to the dungeon dimensions is now fully open.


### Fixed

- Fix escaping of special character in URLs to write_url().

  Issue [#45](https://github.com/jmcnamara/excel-writer-xlsx/issues/45).

- Fix for 0 access/modification date on vbaProject.bin files extracted
  using extract_vba. The date isn't generally set correctly in the
  source xlsm file but this caused issues on Windows.


## [0.59] - 2012-11-26

### Added

- Added macro support via VBA projects extracted from existing Excel
  xlsm files. User defined functions can be called from worksheets
  and macros can be called by the user but they cannot, currently,
  be linked to form elements such as buttons.


## [0.58] - 2012-11-23

### Added

- Added chart area and plot area formatting.


## [0.57] - 2012-11-21

### Added

- Add major and minor axis chart gridline formatting.


## [0.56] - 2012-11-18

### Fixed

- Fix for issue where chart creation order had to be the same
  as the insertion order or charts would be out of sync.
  Frederic Claude Sievert and Hurricup. Issue [#42](https://github.com/jmcnamara/excel-writer-xlsx/issues/42).

- Fixed issue where gridlines didn't work in Scatter and Stock
  charts. Issue [#41](https://github.com/jmcnamara/excel-writer-xlsx/issues/41).

- Fixed default XML encoding to avoid/solve various issues with XML
  encoding created by the XML changes in version 0.51. Issue [#43](https://github.com/jmcnamara/excel-writer-xlsx/issues/43).


## [0.55] - 2012-11-10

### Added

- Added Sparklines.


### Fixed

- Fix for issue with "begins with" and "ends with" Conditional
  Formatting. Issue [#40](https://github.com/jmcnamara/excel-writer-xlsx/issues/40).


## [0.54] - 2012-11-05

### Added

- Added font manipulation to Charts.

- Added number formats to Chart axes.

- Added Radar Charts.


### Fixed

- Fix for XML encoding in write_url() internal/external
  links. Issue [#37](https://github.com/jmcnamara/excel-writer-xlsx/issues/37).


## [0.53] - 2012-10-10

### Fixed

- Fix for broken MANIFEST file.


## [0.52] - 2012-10-09

### Fixed

- Added dependency on Date::Calc to xl_parse_date.t test.
  Closes #30 and RT#79790.

- Fix for XML encoding of URLs. Closes #31.


### Added

- Refactored XMLWriter into a single class. This breaks the last
  remaining ties to XML::Writer to allow for future additions
  and optimizations. Renamed methods for consistency.


## [0.51] - 2012-09-16

### Added

- Speed optimizations.
  This release contains a series of optimizations aimed
  at increasing the speed of Excel::Writer::XLSX. The
  overall improvement is around 66%.
  See the SPEED AND MEMORY USAGE section of the documentation.

- Memory usage optimizations.
  This fixes an issue where the memory used for the worksheet
  data tables was freed but then brought back into usage due
  to the use of an array as the base data structure. This
  meant that the memory usage still continued to grow with
  large row counts.


### Fixed

- Added warning about Excel limit to 65,530 urls per worksheet.

- Limit URLs to Excel's limit of 255 chars. Fixes Issue [#26](https://github.com/jmcnamara/excel-writer-xlsx/issues/26).

- Fix for whitespace in urls. Fixes Issue [#25](https://github.com/jmcnamara/excel-writer-xlsx/issues/25).

- Fix for solid fill of type 'none' is chart series.
  Closes issue #27 reported on Stack Overflow.

- Modified write_array_formula() to apply format over full range.
  Fixes issue #18.

- Fix for issue with chart formula referring to non-existent sheet name.
  It is now a fatal error to specify a chart series formula that
  refers to an non-existent worksheet name. Fixes issue #17.


## [0.50] - 2012-09-09

### Added

- Added option to add secondary axes to charts.
  Thanks to Eric Johnson and to Foxtons for sponsoring the work.

- Added add_table() method to add Excel tables to worksheets.


### Fixed

- Fix for right/left auto shape connection when destination
  is left of source shape. Thanks to Dave Clarke for fix.

- Fix for issue #16. Format::copy() method not protecting values.
  The Format copy() method over-writes certain new properties that
  weren't in Spreadsheet::WriteExcel. This fixes the issue by
  storing and restoring the properties during copy.

- Fix for issue #15: write_url with local sub directory.
  Local sub-directories were incorrectly treated as
  file:// external.

- Fix for for issue #14: Non-numeric data in chart value axes
  are now converted to zero in chart data cache, as required
  by Excel.


## [0.49] - 2012-07-12

### Added

- Added show_blanks_as() chart method to control the display of
  blank data.

- Added show_hidden_data() chart method to control the display of
  data in hidden rows and columns.


### Fixed

- Added fix for fg/bg colors in conditional formats which are
  shared with cell formats.
  Reported by Patryk Kwiatkowski.

- Fix for xl_parse_time() with hours > 24. Github issue #11.

- Fixed lc() warning in Utility.pm in recent perls. Github issue #10.

- Fixed issue with non-integer shape dimensions. Thanks Dave Clarke.

- Fixed error handling for shape connectors. Thanks Dave Clarke.


## [0.48] - 2012-06-25

### Added

- Added worksheet shapes. A major new feature.
  Patch, docs, tests and example programs by Dave Clarke.

- Added stacked and percent_stacked chart subtypes to Area charts.


### Fixed

- Added fix for chart names in embedded charts.
  Reported by Matt Freel.

- Fixed bug with Unicode characters in rich strings.
  Reported by Michiel van Rhee.


## [0.47] - 2012-04-10

### Added

- Additional conditional formatting options such as color, type and value
  for 2_color_scale, 3_color_scale and data_bar. Added option for non-
  contiguous data ranges as well.

- Additional chart data label parameters such as position, leader lines
  and percentage. Initial patch by George E. Tarrant III.


### Fixed

- Fixed for Autofilter filter_column() offset bug reported by
  Krishna Rajendran.

- Fix for write_url() where url contains invalid whitespace, RT #75808,
  reported by Oleg G. The write_url() method now throws a warning and
  rejects the invalid url to avoid file corruption.


## [0.46] - 2012-02-10

### Fixed

- Fix for x-axis major/minor units in scatter charts.
  Reported by Carey Drake.


## [0.45] - 2012-01-09

### Fixed

- Changed from File::Temp tempdir() to newdir() to cleanup the temp dir at
  object destruction  rather than the program exit. Also improved error
  reporting when mkdir() fails.
  Reported by Kevin Ruscoe.

- Fix to escape control characters in strings.
  Reported by Kevin Ruscoe.


## [0.44] - 2012-01-05

### Fixed

- Fix for missing return value from Workbook::close() with filehandles.
  RT 73724. Reported and patched by Charles Bailey.

- Fixed support special filename/filehandle '-'.
  RT 73424. Reported by YuvalL and Charles Bailey.

- Fix for non-working reverse x_axis with Scatter charts.
  Reported by Viqar Abbasi.


## [0.43] - 2011-12-18

### Added

- Added chart axis label position option.

- Added invert_if_negative option for chart series fills.


## [0.42] - 2011-12-17

### Fixed

- Fix for set_optimization() where first row isn't 0.
  Reported by Giulio Orsero.

- Fix to preserve whitespace in inline strings.
  Reported by Giulio Orsero.


## [0.41] - 2011-12-10

### Fixed

- Increased IO::File requirement to 1.14 to prevent taint issues on some
  5.8.8/5.8.6 platforms.


## [0.40] - 2011-12-07

### Fixed

- Fix for unreadable xlsx files when generator program has -l on the
  commandline or had redefined $/. Github issue #7.
  Reported by John Riksten.


## [0.39] - 2011-12-03

### Fixed

- Fix for spurious Mac ._Makefile.PL in the distro which prevented
  automated testing and installation. Github issue #5.
  Reported by Tobias Oetiker.

- Fix for failing test sub_convert_date_time.t due to extra precision
  on longdouble perls. RT #71762
  Reported by Douglas Wilson.


## [0.38] - 2011-12-03

### Added

- Backported from perl 5.10.0 to perl 5.8.2.
  You are killing me guys. Killing me.


## [0.37] - 2011-12-02

### Added

- Added additional axis options: minor and major units, log base
  and axis crossing.


## [0.36] - 2011-11-29

### Added

- Added "min" and "max" options to axis ranges via set_x_axis() and
  set_y_axis.


## [0.35] - 2011-11-27

### Added

- Added Scatter chart subtypes: markers_only (the default),
  straight_with_markers, straight, smooth_with_markers and smooth.


## [0.34] - 2011-11-04

### Added

- Added set_optimization() method to reduce memory usage for very large
  data sets.


## [0.33] - 2011-10-28

### Added

- Added addition conditional formatting types: cell, date, time_period,
  text, average, duplicate, unique, top, bottom, blanks, no_blanks,
  errors, no_errors, 2_color_scale, 3_color_scale, data_bar  and formula.


## [0.32] - 2011-10-20

### Fixed

- Fix for format alignment bug.
  Reported by Roderich Schupp.


## [0.31] - 2011-10-18

### Added

- Added basic conditional formatting via the conditional_format()
  Worksheet method. More conditional formatting types will follow.

- Added conditional_format.pl example program.


## [0.30] - 2011-10-06

### Added

- Added stacked and percent_stacked chart subtypes to Bar and Column
  chart types.


## [0.29] - 2011-10-05

### Added

- Added the merge_range_type() method for finer control over the types
  written using merge_range().


## [0.28] - 2011-10-04

### Added

- Added default write_formula() value for compatibility with Google docs.

- Updated Example.pm docs with Excel 2007 images.


## [0.27] - 2011-10-02

### Added

- Excel::Writer::XLSX is now 100% functionally and API compatible
  with Spreadsheet::WriteExcel.

- Added outlines and grouping functionality.

- Added outline.pl and outline_collapsed.pl example programs.


## [0.26] - 2011-10-01

### Added

- Added cell comment methods and options.
  Thanks to Barry Downes for providing the interim functionality

- Added comments1.pl and comments2.pl example programs.


## [0.25] - 2011-06-16

### Added

- Added option to add defined names to workbooks and worksheets.
  Added defined_name.pl example program.


### Fixed

- Fix for fit_to_pages() with zero values.
  Reported by Aki Huttunen.


## [0.24] - 2011-06-11

### Added

- Added data validation and data_validate.pl example.

- Added the option to turn off data series in chart legends.


## [0.23] - 2011-05-26

### Fixed

- Fix for charts ranges containing empty values.


## [0.22] - 2011-05-22

### Added

- Added 'reverse' option to set_x_axis() and set_y_axis() in
  charts.


## [0.21] - 2011-05-11

### Fixed

- Fixed support for filehandles.


### Added

- Added write_to_scalar.pl and filehandle.pl example programs.


## [0.20] - 2011-05-10

### Fixed

- Fix for programs running under taint mode.


### Added

- Added set_tempdir().


### Fixed

- Fix for color formatting in chartsheets.


## [0.19] - 2011-05-05

### Added

- Added new chart formatting options for line properties,
  markers, trendlines and data labels. See Chart.pm.

- Added partial support for insert_image().

- Improved backward compatibility for deprecated methods
  store_formula() and repeat_formula().


### Fixed

- Fixed missing formatting for array formulas.
  Reported by Cyrille Gourves.

- Fixed issue with chart scaling that caused "unreadable content"
  Excel error.


## [0.18] - 2011-04-07

### Added

- Added set_properties() method to add document properties.
  Added properties.pl and tests.


## [0.17] - 2011-04-04

### Added

- Added charting feature. See Chart.pm.


### Fixed

- Fix for file corruption issue when there are more than 10 custom colors.
  Reported by Brian R. Landy.


## [0.16] - 2011-03-04

### Fixed

- Clarified support for deprecated methods in documentation and added
  backward compatible methods in some cases.

- Fix for center_horizontally() issue.
  Reported by Giulio Orsero.

- Fix for number like strings getting written as strings instead of numbers.
  Reported by Giulio Orsero.


## [0.15] - 2011-03-01

### Fixed

- Fix for issues with set_row() not passing on format to cells
  in the row. Reported by Giulio Orsero.

- Fixes for related issue in set_column().


## [0.14] - 2011-02-26

### Added

- Added write_rich_string() method to write a string with multiple
  formats.

- Added rich_strings.pl example program.

- Added set_1904() method for dates with a 1904 epoch.

- Added date_time.pl example program.


### Fixed

- Fixed issue where leading and trailing whitespace in cell strings
  wasn't preserved.


## [0.13] - 2011-02-22

### Added

- Added additional page setup methods:
  set_zoom()
  right_to_left()
  hide_zero()
  set_custom_color()
  set_tab_color()
  protect()

- Added Cell property methods:
  set_locked()
  set_hidden()

- Added example programs:
  hide_sheet.pl
  protection.pl
  right_to_left.pl
  tab_colors.pl


## [0.12] - 2011-02-19

### Added

- Added set_selection() method for selecting cells.


## [0.11] - 2011-02-17

### Fixed

- Fix for temp dirs not been removed after xlsx file creation.
  http://rt.cpan.org/Ticket/Display.html?id=65816
  Reported by Andreas Koenig.


## [0.10] - 2011-02-17

### Added

- Added freeze_panes() and split_panes().

- Added panes.pl example program.


## [0.09] - 2011-02-13

### Added

- Added write_url() for internal and external hyperlinks.

- Added hyperlink1+2.pl example programs.


## [0.08] - 2011-02-03

### Added

- Added autofilter(), column_filter() and column_filter_list() methods.

- Added autofilter.pl example program.


## [0.07] - 2011-01-28

### Added

- Added additional Page Setup methods.
  set_page_view()
  repeat_rows()
  repeat_columns()
  hide_gridlines()
  print_row_col_headers()
  print_area()
  print_across()
  fit_to_pages()
  set_start_page()
  set_print_scale()
  set_h_pagebreaks()
  set_v_pagebreaks()

- Added headers.pl example program.


## [0.06] - 2011-01-19

### Fixed

- Added fix for XML characters in attributes.
  Reported by John Roll.


### Added

- Added initial Page Setup methods.
  set_landscape()
  set_portrait()
  set_paper()
  center_horizontally()
  center_vertically()
  set_margins()
  set_header()
  set_footer()


## [0.05] - 2011-01-04

### Added

- Added support for array_formulas. See the docs for write_array_formula()
  and the example program.


## [0.04] - 2011-01-03

### Added

- Added merge_range() for merging cells. With tests and examples.


## [0.03] - 2011-01-03

### Added

- Optimizations. The module is now 100% faster.


## [0.02] - 2010-10-12

### Fixed

- Fixed dependencies in Makefile.


## [0.01] - 2010-10-11

### Added

- First CPAN release.

