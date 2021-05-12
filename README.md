# MSTemplate_Creator

<!-- badges: start -->

[![License:
MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/SLINGhub/MSTemplate_Creator/blob/master/LICENSE.md)
<!-- badges: end -->

MSTemplate_Creator is an excel macro file created to provide users a
more friendly interface to take in MRM transition names data exported
directly from mass spectrometry software to create several annotation
templates suited for automated data processing and statistical analysis.

It is currently distributed as platform independent source code under
the MIT license.

## Starting Up

Download the repository and open the excel macro file
`MSTemplate_Creator.xlsm`

![OpenMSTemplate_Creator](figures/OpenMSTemplate_Creator.JPG)

Upon opening you may encounter this security warning. Click on **Enable
Content** so that the macro in the file will be activated.

![EnableContent](figures/EnableContentWarning.jpg)

## Using Transition_Annot Sheet

Load transition names from Agilent MRM data in csv file with **Load
Transition_Name from Raw Data**

![Load_Transition_Name_from_Raw_Data](figures/Load_Transition_Name_from_Raw_Data.gif)

Load transition names from tabular data in csv file with **Load
Transition_Name from Table Data**

![Load_Transition_Name_from_Table_Data](figures/Load_Transition_Name_from_Table_Data.gif)

Check the internal standards with **Validate ISTD**.

![Validate_ISTD](figures/Validate_ISTD.gif)

Once validated, transfer the internal standards to sheet `ISTD_Annot`
with **Load ISTD to ISTD_Table**

![Load_ISTD_to_ISTD_Table](figures/Load_ISTD_to_ISTD_Table.gif)

## Using ISTD_Annot Sheet

Key in the concentration of the internal standard and convert to nM or
other units to verify. Unit values under the column `Custom_Unit` can be
used later to obtain the sample unit of concentration.

![Convert_to_nM](figures/Convert_to_nM.gif)

## Using Sample_Annot Sheet

Load sample names from Agilent MRM data in csv file with **Load Sample
Annotation from Raw Data**. Use **Autofill ‘Sample’ in Sample_Type** to
fill empty cells under the `Sample_Type` column with “SPL”

![Load_Sample_Annotation_from_Raw_Data](figures/Load_Sample_Annotation_from_Raw_Data.gif)

It is possible to merge Agilent MRM data with a sample annotation file
in csv.

![Merge_Raw_Data_with_Sample_Annotation](figures/Merge_Raw_Data_with_Sample_Annotation.gif)

Load sample names from tabular data in csv file with **Load Sample
Annotation from Table Data**. Use **Autofill ‘Sample’ in Sample_Type**
to fill empty cells under the `Sample_Type` column with “SPL”

![Load_Sample_Annotation_from_Table_Data](figures/Load_Sample_Annotation_from_Table_Data.gif)

Next, fill in the `Sample_Amount`, `Sample_Amount_Unit` and the
`ISTD_Mixture_Volume_[uL]` columns. If a particular `Sample_Type` has
consistent inputs, the **Autofill by Sample_Type** button helps to fill
in these consistent values quickly.

![Autofill_by_Sample_Type](figures/Autofill_by_Sample_Type.gif)

To obtain the analyte’s concentration unit measured in each sample, go
the `Sample_Annot` sheet and fill in the `Sample_Amount_Unit` for each
sample. Next, on the `ISTD_Annot` sheet, select the concentration unit
of the internal standard to use under the `Custom_Unit` column. Return
to the `Sample_Annot` sheet and use **Autofill Concentration_Unit** to
fill in the `Concentration_Unit` column.

![Autofill_Concentration_Unit](figures/Autofill_Concentration_Unit.gif)

Transfer Sample with QC sample type of “RQC” to `Dilution_Annot` sheet
with **Load RQC Samples to Dilution_Table**

![Load_RQC_Samples_to_Dilution_Table](figures/Load_RQC_Samples_to_Dilution_Table.gif)
