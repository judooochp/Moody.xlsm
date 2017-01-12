# MoodyCanvas
Wireframe Representations of Surface Plate Flatness Measurements using the Moody Method in MSExcel

System Requirements:
Microsoft Excel 2016 preferred, other versions may have formatting or functional inconsistencies.
Macros must be enabled.
MS Internet Explorer 9 or better. Other HTML5 (and canvas) capable browsers are fine, but won't automatically open.

Originally created in MSExcel 2016, this file is an implementation of T.O. 33K6-4-137-1: Surface Plates Autocollimator Method, dated 30 October 2005. It contains eight procedural sheets for entry of arcsecond angle measurements from which are calculated elevations, one sheet for setup entry and a results summary, and one hidden sheet that replicates a table from the T.O. for easy =IF() statements in the Setup&Summary sheet.
A VBA macro takes the results and generates an HTML file that implements two <canvas> elements. One is a 3D representation of the surface plate, and the other is a point-by-point integer map of the elevation at each measurement point.
Created by Jacob Gilliam, January2017
