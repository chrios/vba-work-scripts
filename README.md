# Visual Basic for Applications Work Scripts

This repo holds a script I created to solve a common problem in industrial engraving. Labels require holes placed precisely so that screws can be used as fixings. There is usually a standard for this. CorelDRAW does not have the facility to easily and quickly place these holes - a common feature of industrial engraving software. I implemented a script that will insert the circle objects for the laser to cut out. There is a form to capture user input for hole the distance from edges horizontally, vertically, the diameter of the holes, and the cutoff for one hole on each side vs one hole in each corner.

## To install

Open the Visual Basic Macro editor in CorelDRAW. Create a new Macro, and copy the contents of screwHoles.vbs. Create a new form, create four fields named:

  formXDistance
  formYDistance
  formDiameter
  formCutOff  

Edit the callback code for the form and paste the contents of frmHoles.vbs in.

Run the macro. Don't forget to add a submit button to the form.
