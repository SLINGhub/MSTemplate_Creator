# MSTemplate\_Creator

MSTemplate\_Creator is an excel macro file created to provide users a
more friendly interface to take in MRM transition names data exported
directly from mass spectrometry software to create several annotation
templates suited for automated data processing and statistical analysis.

It is currently distributed as platform independent source code under
the MIT license.

## Starting Up

Download the repository and open the excel macro file
`MSTemplate_Creator.xlsm`

![OpenMSTemplate\_Creator](figures/OpenMSTemplate_Creator.JPG)

Upon opening you may encounter this security warning. Click on **Enable
Content** so that the macro in the file will be activated.

![EnableContent](figures/EnableContentWarning.jpg)

## Using Transition\_Annot Sheet

Load Agilent MRM data in csv file with **Load Transition\_Name from Raw
Data**

![Load Transition\_Name from Raw Data](figures/trial.gif)

Load tabular data in csv file with **Load Transition\_Name from Table
Data**

![Load Transition\_Name from Table Data](figures/trial2.gif)

Check the internal standards with **Validate ISTD**

![Validate ISTD](figures/trial3.gif) Once validated, transfer the
internal standards to sheet `ISTD_Annot` with **Load ISTD to
ISTD\_Table**
