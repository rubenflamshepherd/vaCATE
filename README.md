Visualized Automator of Compartment Analysis by Tracer Efflux (vaCATE)
============

Python program for compartmental analysis by tracer efflux (CATE).

To analyze CATE data, data sets must be influx into a template file. The template file can be generated from the initial window presented by vaCATE by pressing the "Generate CATE Template" button.

Template Rules:
1) Every bold field must be contain information (i.e., rows 1-7 as labeled in columns 1-2).
3) Do not provide titles for an empty column.
4) For Elution time (Column under cell B8), seconds should be expressed as fractions of a minute (e.g., 30 sec = 0.5).
5) The 'Vial #" Field is optional.

Preconstructed template files for demostration purposes are available in the 'Tests' and 'Examples' subfolders in the folder that vaCATE was in installed/cloned into.

Using the installer is recommended over cloning the repository. This is because vaCATE was initially created using wxPython 2.8.12.1, which is no longer available. Because of this, a local version of wxPython 3.03 is provided in the 'local' folder when using 'requirements.txt' to install dependencies. However, this results in some spacing issues in the initial window presented by vaCATE, which can be solved by simply resizing the window.

For inquires please contact me either here on github or at ruben dot flam dot shepherd at gmail dot com.
