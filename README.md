# MSTemplate_Creator

<img src="figures/logo.png" align="right" height="200" />

<!-- badges: start -->

[![License:
MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/SLINGhub/MSTemplate_Creator/blob/main/LICENSE.md)
[![Maintainer:
JauntyJJS](https://img.shields.io/badge/Maintainer-JauntyJJS-blue.svg)](https://github.com/JauntyJJS)
[![Excel](https://img.shields.io/badge/Excel-2019%2064%20bit-brightgreen)](https://www.microsoft.com/en-sg/microsoft-365/excel)
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.6552782.svg)](https://doi.org/10.5281/zenodo.6552782)
<!-- badges: end -->

`MSTemplate_Creator` is an excel template providing users with a
friendly interface to create different metadata
annotation templates from mass spectrometry analysis results, which are suited for subsequent automated data processing and
statistical analysis. To ensure integrity of the metdata, `MSTemplate_Creator` allows to read in MRM data files exported
from mass spectrometry software, and provides macros to valide the metadata.

<br><br><br>
![AboutMSTemplate_Creator](figures/AboutMSTemplate_Creator.png)

# Table of Content

- [MSTemplate_Creator](#mstemplate_creator)
- [Table of Content](#table-of-content)
- [Meta](#meta)
- [Overview](#overview)
- [Starting Up](#starting-up)
- [Using Transition_Annot Sheet](#using-transition_annot-sheet)
- [Using ISTD_Annot Sheet](#using-istd_annot-sheet)
- [Using Sample_Annot Sheet](#using-sample_annot-sheet)

# Meta

- We welcome contributions from general questions to bug reports. Check
  out the [contributions](CONTRIBUTING.md) guidelines. Please note that
  this project is released with a [Contributor Code of
  Conduct](https://www.contributor-covenant.org/version/2/0/code_of_conduct/).
  By participating in this project you agree to abide by its terms.
- License:
  [MIT](https://github.com/SLINGhub/MSTemplate_Creator/blob/main/LICENSE.md)
- Think `MSTemplate_Creator` is useful? Let others discover it, by
  telling them in person, via Twitter
  [![Tweet](https://img.shields.io/twitter/url/http/shields.io.svg?style=social)](https://twitter.com/LOGIN)
  or a blog post. Do use the `🙌 Show and tell` under the [GitHub
  Discussions](https://github.com/SLINGhub/MSTemplate_Creator/discussions)
  and give this repository a star as well.

![GitHubStar](figures/GitHubStar.JPG)

- If you wish to acknowledge the use of this software in a journal
  paper, please include the version number. For reproducibility, it is
  advisable to use the software from the
  [Releases](https://github.com/SLINGhub/MSTemplate_Creator/releases)
  section in GitHub rather than from the main branch.
- To date, the software is only able to take in the following files
  exported from the following software.
  - csv files from [Agilent MassHunter Quantitative Analysis
    Software](https://www.agilent.com/en/product/software-informatics/mass-spectrometry-software/data-analysis/quantitative-analysis)
- Refer to the [NEWS.md
  file](https://github.com/SLINGhub/MSTemplate_Creator/blob/main/NEWS.md)
  to see what is being worked on as well as update to changes between
  back to back versions.

[Back to
top](https://github.com/SLINGhub/MSTemplate_Creator#mstemplate_creator)

# Overview

For an overview on how the tool works, take a look at the Summary and
Familiarisation file in the
[documentation](https://github.com/SLINGhub/MSTemplate_Creator/tree/main/docs)
page.

![SummaryCheatSheet](figures/SummaryCheatSheet1.JPG)

![SummaryCheatSheet](figures/SummaryCheatSheet2.JPG)

[Back to
top](https://github.com/SLINGhub/MSTemplate_Creator#mstemplate_creator)

# Starting Up

Go to the
[Releases](https://github.com/SLINGhub/MSTemplate_Creator/releases)
section in GitHub.

![ReleasesWebpage](figures/ReleasesWebpage.JPG)

Download the zip folder. Unzip the folder, double click on the excel
macro file `MSTemplate_Creator.xlsm` to start.

![OpenMSTemplate_Creator](figures/OpenMSTemplate_Creator.JPG)

Upon opening you may encounter this security warning. Click on **Enable
Content** so that the macro in the file will be activated.

![EnableContent](figures/EnableContentWarning.jpg)

[Back to
top](https://github.com/SLINGhub/MSTemplate_Creator#mstemplate_creator)

# Using Transition_Annot Sheet

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

[Back to
top](https://github.com/SLINGhub/MSTemplate_Creator#mstemplate_creator)

# Using ISTD_Annot Sheet

Key in the concentration of the internal standard and convert to nM or
other units to verify. Unit values under the column `Custom_Unit` can be
used later to obtain the sample unit of concentration.

![Convert_to_nM](figures/Convert_to_nM.gif)

[Back to
top](https://github.com/SLINGhub/MSTemplate_Creator#mstemplate_creator)

# Using Sample_Annot Sheet

Load sample names from Agilent MRM data in csv file with **Load Sample
Annotation from Raw Data**.

Use **Autofill ‘Sample’ in Sample_Type** to fill empty cells under the
`Sample_Type` column with “SPL”

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

[Back to
top](https://github.com/SLINGhub/MSTemplate_Creator#mstemplate_creator)
