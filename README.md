# MoodyCanvas
Wireframe Representations of Surface Plate Flatness Measurements using a Modified Moody Method in MSExcel to Export HTML Canvas Printables
_______________________________________________________________________________________________

Intended Updates:

(Completed 14Jan2017) Place HTML textbox and button that changes angle of view (rotate Z axis)
  
(Completed 13Jan2017) Organize the different sheets in a way that makes sense to more than just Jake

  -Split Setup and Summary; place Summary after Line H

(Completed 23Oct2017) Test Files for several sizes

Diffuse Closure error across Lines G and H to clean up graphic.
  -Make centerpoint Zero
  -Still report closure error on header
  
Label graphic's Lines <strike>A-H</strike> 1-8

Make compatible with former versions of Excel (97-2003)

Handle Customer Requirements/OOT conditions

(Partially Complete 15Jan2017) Large plates with large elevations: second Canvas numbers run over each other, especially on horizontal lines. Fix it.

  -(Needs Testing)-- Defeat staggering for horizontal lines of smaller plates.
  -(15Jan2017)-- Expanded X dimension with ratio of (#diagonalMeasurements/#horixontalMeasurements)x0.85 on the multiplier for X axis values

(Completed 28Jan2017) Make extensible to plates that are not standard sizes using the GGG-P-463C formula.

(Completed 28Nov2017) Open HTML automatically. Now opens in user's default browser.
_______________________________________________________________________________________________

System Requirements:

Microsoft Excel 2007 or later preferred, other versions may have formatting or functional inconsistencies.

Macros must be enabled.

MS Internet Explorer 9 or better. Other HTML5 (and canvas) capable browsers are fine, but won't automatically open.

Originally created in MSExcel 2016, this file is an implementation of T.O. 33K6-4-137-1: Surface Plates Autocollimator Method, dated 30 October 2005. It contains eight procedural sheets for entry of arcsecond angle measurements from which are calculated elevations, one sheet for setup entry and a results summary, and one hidden sheet that replicates a table from the T.O. for easy =IF() statements in the Setup&Summary sheet.

Procedure calls for step to bring Lines G and H centerpoint to Zero with the Lines A and B. I have skipped this to show closure error in graphic.

Changes to Setup dynamically makes available required number of rows on Line sheets to properly execute math for Summary.

A VBA macro takes the results and generates an HTML file that implements two <canvas> elements. One is a 3D representation of the surface plate, and the other is a point-by-point integer map of the elevation at each measurement point.
Created by Jacob Gilliam, January2017
