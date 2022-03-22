# MSTemplate_Creator 1.1.0.9000 (development version)

## TODO

* Check if relative sample amount definition is clear enough.
* Decide on the most suited open source licence.

## Done
* Try out the code inspection feature in [RubberDuck](http://rubberduckvba.com/).
* Add documentation for the Visual Basic functions in the excel sheet.
* Give instructions on how to use Natural Docs in Developers documentation and readme guide

# MSTemplate_Creator 1.1.0

* Fix bugs when reading transition names with qualifiers in Agilent Compound Table form.
* Change concentration unit to have uL as the denominator instead of mL.
* Create unit test for reading transition names with Qualifiers for Agilent Wide Table and Compound Table form.
* Create unit test for finding concentration unit.
* Update documentation on the concentration unit test.
* Add an integration test "Nothing_To_Transfer_Test" to check that the program can still run when there is no data to another sheet.


# MSTemplate_Creator 1.0.1

* Changed dilution factor to relative sample amount
* Added a logo

# MSTemplate_Creator 1.0.0

* Added some Github related markdown files like issue templates, contributing guidelines and code of conduct.

# MSTemplate_Creator 0.0.2

* Transfer respository from Bitbucket to GitHub
* Update the Sample Type to be the same as the LIMS in SLING.
* Changed the excel sheet font to "Consolas" so that the number "0" and the letter "O" can be differentiated easily.
* Match the buttons with the relevant column colours in the sheet.
* Added a new button to auto fill the sample amount, sample amount unit and the istd mixture volumne column based on the QC sample type.
* If user change the concentration Custom_Unit in the ISTD_Annot sheet and there are values in the Custom Unit in the ISTD_Annot sheet and values in the Sample Amount in the Sample_Annot sheet, the software will auto fill the concentration unit. This is to ensure the right concentration unit is updated correctly.

# MSTemplate_Creator 0.0.1

* Added a `NEWS.md` file to track changes to the package.
* Aim to create a git tag version.
