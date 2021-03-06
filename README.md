<h1 align='center'>Visualized Automator of<br> Compartmental Analysis by Tracer Efflux (vaCATE)</h1>

<h3 align='center'>
<a href='https://openresearchsoftware.metajnl.com/articles/10.5334/jors.175/'>
As published in the Journal of Open Research Software
</a>
</h3>

<b>Abstract:</b> Compartmental analysis by tracer efflux (CATE) is fundamental to examinations of membrane transport, allowing study of solute movement among subcellular compartments with high temporal, spatial, and chemical resolution. CATE can provide a wealth of information about fluxes and pool sizes in complex systems, but is a mathematically intensive procedure, and there is a need for software designed to fully, easily, and dynamically analyse results from CATE experiments. Here we present vaCATE (Visualized Automation of Compartmental Analysis by Tracer Efflux), a software package that meets these criteria. A robust suite of test cases using CATE datasets from experiments with intact rice (Oryza sativa L.) root systems reveals the high fidelity of vaCATE and the ease with which parameters can be extracted, using a three-compartment model and a curve-stripping procedure to distinguish them on the basis of variable exchange rates. vaCATE was developed using Python 2.7 and can be used in most situations where compartmental analysis is required.

___

<h3 align='center'>Basic Operation</h3>

vaCATE is a Python program designed to automate and maximize data extraction from compartmental analysis by tracer efflux (CATE).

To analyze CATE data, data sets must be input into a template file. The template file can be generated from the initial window presented by vaCATE by pressing the "Generate CATE Template" button (Figure 1 below).

<p align="center">
<img src="https://github.com/rubenflamshepherd/vaCATE/blob/master/Images/Figure%201.png">
<br><b>Figure 1. </b>Initial window presented by vaCATE.<br>
</p>

Preconstructed template files for demostration purposes are available in the 'Tests' and 'Examples' subfolders in the folder that vaCATE was installed/cloned into. See Figure 2 below for an example.

<p align="center">
<img width=300 src="https://github.com/rubenflamshepherd/vaCATE/blob/master/Images/Figure%202.png">
<br><b>Figure 2. </b>Partial example of a filled-in template file.<br>
</p>

Template Rules:
1) Every bold field must contain information (i.e., rows 1-7 as labeled in columns 1-2).
3) Do not provide titles for an empty column.
4) For Elution time (Column under cell B8), seconds should be expressed as fractions of a minute (e.g., 30 sec = 0.5).
5) The 'Vial #" Field is optional.

Once a properly filled in template is input to vaCATE, the analysis can be dyanmically previewed before export into an Excel file (see Figure 3 below)

<p align="center">
<img src="https://github.com/rubenflamshepherd/vaCATE/blob/master/Images/Figure%204.png">
<br><b>Figure 3. </b>Dynamic preview of analysis.<br>
</p>

Using the installer ([available here](https://doi.org/10.6084/m9.figshare.4688503.v2)) is recommended over cloning the repository. This is because vaCATE was initially created using wxPython 2.8.12.1, which is no longer available on PyPI (Python Package Index) or elsewhere. Because of this, when cloning the repository and using 'requirements.txt' to install dependencies, wxPython 3.03 is provided in the 'local' folder. However, versioning differences result in some spacing issues in the initial window presented by vaCATE, which can be solved by simply resizing the window.

For inquires please contact me either here on github or at ruben dot flam dot shepherd at gmail dot com.

<h2 align='center'>Supplemental Information to Published Paper</h2>

<h3 align='center'>Accounting for Shifting Data Series’</h3>

The data-validation step in the curve-stripping function poses some interesting challenges from a software engineering perspective. Most poignantly, data points used as boundaries for earlier, more rapidly exchanging phases may be removed. Additionally, a sufficient number of data points may be removed such that compartmental analysis is rendered impossible (i.e., the boundaries delineating a phase contain less than two data points).

To account for these possibilities, several steps are taken. Firstly, the series of elution time points from which negative log operations have potentially been removed (see above) is stored separately (in the ‘elut_ends_parsed’ Analysis attribute) from the original elution series (which is stored in the ‘elut_ends’ attribute). This strict separation makes it explicit when calculations are using a data series that may be missing expected points. Secondly, as the range and contents of the data series being used can vary depending on the phase being examined and the aforementioned removal of data points, indices of data series are not used outside of local scopes. Instead, explicit elution time points are used for phase boundaries. Whenever an index becomes required, the relevant elution time point is converted to an index that is specific to the data series in question. This is done by the x_to_index() method in the Operations module.

The x_to_index() method is itself is one of the more important mechanisms through which errors are avoided. Given an elution time point, the type of boundary it represents, the local x-series of interest (constructed from ‘elut_ends_parsed’), and the larger, unparsed x-series (‘elut_ends’), x_to_index() first checks if this elution time point has been removed from the local x-series. If it hasn’t, returning the index in the local x-series at which it occurs is straightforward. However, if it has been removed, a different elution time point must be used so that the returned index corresponds to a point that exists in the local x-series. To find this index, a neighbouring elution time point is chosen from the larger x-series. To ensure the new elution time point being used is in the proper range, the boundary type of the original elution time point is considered. If the start of a phase boundary was represented, then the following elution time point in the larger x-series is chosen. Conversely, if the end of a phase boundary was represented, then the preceding elution time point in the larger x-series is chosen. This process is repeated until an elution time point that exists in our local x-series is chosen. The conversion of this point to an index is then straightforward.

<p align="center">
<img src="https://github.com/rubenflamshepherd/vaCATE/blob/master/Images/Figure%20S1.png">
<br><b>Figure S1. </b>The extract_phase() method in the Operations module is used to return a Phase object given a set of boundaries (xs) and properly curve-stripped data-series’ (x_series and y_series).<br>
</p>

The extract_phase() method also contains several error-checking mechanisms that work in conjunction with x_to_index(). To this end, the first thing done in the method is the confirmation that the data series to be analysed is of sufficient length (>1) for linear regression (Fig. S1; lines 174-181). If this check is failed (as would be the case if sufficient points have been removed in the parsing process) then compartmental analysis of the phase in question is not completed and a blank Phase object is returned. After this data-validation point, the elution points being passed as the boundaries for the phase in question (xs) are converted to indices corresponding to the local data series (x_series; Fig. S1; lines 183-189). These indices are then used to take list slices (x_phase and y_phase) of the local data series (Fig. 6; lines 191-192), demarcating the data series that will be used for the compartmental analysis. In the given description of x_to_index() (see above), it is possible for an elution time point representing the beginning of a phase boundary to be converted to an index that corresponds to an elution point that occurs later than the end of the phase (and vice versa for an elution time point that demarcates the end of a phase boundary). Because these indices are used to slice a list, this does not result in an error, as a list slice containing a start index occurring after an end index simply returns an empty list. The length of these list slices is then also checked, and compartmental analysis is only conducted if the length is sufficient (i.e., >1; Fig. S1; lines 194-206). If this check is failed, a blank Phase object is returned.

<h3 align='center'>Screening User Input for Errors</h3>

As users are allowed to enter the parameters that are to be used for either subjective or objective analysis, the validity of this input must necessarily be screened to avoid unexpected bugs.  Screening is done by the check_obj_input() and check_subj_input() methods of the MainFrame class inside the Preview module (lines 663-685 and 827-874 respectively). Both of these methods implement the RegError class in the same module, which presents the user with a small dialog box containing a meaningful error message that provides feedback as to how the current user input is invalid.

For objective regressions, check_obj_input() limits user input to integers between (inclusive) three and half of the length of the series (rounded down). For subjective regressions, check_subj_input() processes user input through three filters (Fig. 7). First, user input is determined to be either an elution time point or a blank value (“”). Then, the validity of the pairs of data being used to construct individual phase boundaries is determined. That is, both points must be either empty strings or elution time points, with the start of the phase occurring before the end of the phase. Finally, the validity of the phases relative to each other is determined. Specifically, boundaries of phases cannot overlap and boundaries for earlier, more rapidly-exchanging phases can only be defined if boundaries of later, more slowly-exchanging phases are also defined (i.e., boundaries of both Phase III and II must be defined if extraction of parameters from Phase I is required). 

<p align="center">
<img src="https://github.com/rubenflamshepherd/vaCATE/blob/master/Images/Figure%20S2.png">
<br><b>Figure S2. </b>Sequential filters used to determine the validity of<br> user-entered phase boundaries 
for subjective regression in vaCATE.<br>
</p>
