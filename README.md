# MoodyCanvas
Wireframe Representations of Surface Plate Flatness Measurements using the Moody Method in MSExcel
_______________________________________________________________________________________________

Intended Updates:

(Completed 14Jan2017) Place HTML textbox and button that changes angle of view (rotate Z axis)

-Requires exporting x,y,z coordinates and an entire Javascript function to do all the math in the browser, culminating in a clear() and redraw()
  
(Completed 13Jan2017) Organize the different sheets in a way that makes sense to more than just Jake

-Split Setup and Summary; place Summary after Line H

Test Files for several sizes

Diffuse Closure error across Lines G and H to clean up graphic.
  -Make centerpoint Zero
  -Still report closure error on header
  
Label graphic's Lines A-H

Make compatible with former versions of Excel (97-2003)

Handle Customer Requirements/OOT conditions

Large plates with large elevations: second Canvas numbers run over each other, especially on horizontal lines. Fix it.

Make extensible to plates that are not standard sizes using the GGG-P-463C formula.
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
