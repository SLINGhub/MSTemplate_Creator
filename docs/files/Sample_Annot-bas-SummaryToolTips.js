﻿NDSummary.OnToolTipsLoaded("File:Sample_Annot.bas",{144:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype144\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Sub Autofill_Column_By_QC_Sample_Type(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByRef Sample_Type&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Header_Name&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal Autofill_Value&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String</td></tr></table></td><td class=\"PAfterParameters\">)</td></tr></table></div></div><div class=\"TTSummary\">Fill in the column indicated by Header_Name with the value indicated by Autofill_Value on rows whose sample type matches Sample_Type</div></div>",145:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype145\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Sub Create_New_Sample_Annot_Tidy(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal TidyDataFiles&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal DataFileType&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal SampleProperty&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal StartingRowNum&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Integer,</td></tr><tr><td class=\"PModifierQualifier first\">ByVal StartingColumnNum&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">Integer</td></tr></table></td><td class=\"PAfterParameters\">)</td></tr></table></div></div><div class=\"TTSummary\">Create Sample Annotation from an input data file in tabular form and output them into the Sample_Annot sheet. The columns filled will be</div></div>",146:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype146\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Sub Create_New_Sample_Annot_Raw(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal RawDataFiles&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String</td></tr></table></td><td class=\"PAfterParameters\">)</td></tr></table></div></div><div class=\"TTSummary\">Create Sample Annotation from an input raw data file and output them into the Sample_Annot sheet. The columns filled will be</div></div>",147:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype147\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Sub Merge_With_Sample_Annot(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByVal RawDataFiles&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByRef SampleAnnotFile&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String</td></tr></table></td><td class=\"PAfterParameters\">)</td></tr></table></div></div><div class=\"TTSummary\">Merge the an input raw data file with a user input sample annotation file. The merged data is then outputted into the Sample_Annot sheet. The columns filled will be</div></div>",148:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype148\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Get_Sample_Name_Array_From_Annot_File(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByRef xFileName&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String</td></tr></table></td><td class=\"PAfterParameters\">) As String()</td></tr></table></div></div><div class=\"TTSummary\">Get an array of Sample Names from a given sample annotation file in csv and tabular form.</div></div>",149:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype149\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Function Get_Sample_Column_Name_Position_From_Annot_File(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByRef first_line()&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String</td></tr></table></td><td class=\"PAfterParameters\">) As Integer</td></tr></table></div></div><div class=\"TTSummary\">Get the column position where the &quot;Sample Name&quot; column is located as indicated in the Sample_Name text box.</div></div>",150:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype150\" class=\"NDPrototype\"><div class=\"PSection PPlainSection\">Public Function Get_Sample_Annot_Starting_Line_From_Annot_File() As Integer</div></div><div class=\"TTSummary\">Get the starting line where the data is from the annotation file.&nbsp; It should be 0 if the data has no headers and 1 if there is.&nbsp; We assume that the column names is on the first line.&nbsp; Basically, it just check if this check box is checked or not</div></div>",151:"<div class=\"NDToolTip TFunction LVisualBasic\"><div id=\"NDPrototype151\" class=\"NDPrototype WideForm\"><div class=\"PSection PParameterSection CStyle\"><table><tr><td class=\"PBeforeParameters\">Public Sub Load_Sample_Info_To_Excel(</td><td class=\"PParametersParentCell\"><table class=\"PParameters\"><tr><td class=\"PModifierQualifier first\">ByRef xFileName&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String,</td></tr><tr><td class=\"PModifierQualifier first\">ByRef MatchingIndexArray()&nbsp;</td><td class=\"PType\">As&nbsp;</td><td class=\"PName last\">String</td></tr></table></td><td class=\"PAfterParameters\">)</td></tr></table></div></div><div class=\"TTSummary\">Output the sample information (Sample_Amount and ISTD_Mixture_Volume_[uL]) found in the sample annotation file to the Sample_Annot sheet</div></div>"});