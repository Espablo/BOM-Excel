# BOM->Excel
Creates a bill of material and cut lists from the browser components tree in Autodesk Fusion360. Created based on [CSV-BOM](https://github.com/macmanpb/CSV-BOM) 

## General Usage Instructions
After [installation](#installation), go to the toolbar "Create" submenu and choose "Create BOM->Excel". A dialog appears which shows provided options to control the Excel output. Click OK and a save file dialog comes up. Name your file and click OK. After creating the file a message box popups with the information that the file has successfully created. Open a suitable app which can handle Excel formatted files. Import the Excel file and voila the BOM of your design is showing.

### Supportet options

![](resources/BOM-Excel/store_screen.png)

* **Selected only**
> Means that only selected components will be exported to Excel.

* **Exclude if no bodies**
> Components without a body makes no sense. Activate this option to ignore them.

* **Exclude linked**
> If linked components have there own BOM, you can exclude them to keep your BOM lean and clean.

* **Ignore visible state**
> The component is not visible but it should taken to the BOM? Ok, activate this option to do that.

* **Exclude "_"**
> Often users sign components with an underscore to make them visually for internal use. This option ignores such signed components.
> If you deselect this option another option comes up which is descripted next.

* **Strip "_"**
> You want underscore signed components too? No problem, but you dont want the underscore in the outputted component name? Then this option is right for you. It strippes the underscore away.

* **Sort Dimensions**
> If you are cutting your parts out of any kind of panelized material (i.e. plywood), you want the height of the part usually be the thickness of your material. 
If you select this option, the dimensions are sorted to accommodate this, no matter how your part is oriented in your model. 
The smallest value becomes the height (thickness), the next larger the width and the largest the length.

### File

* **Open file**
> Automatically opens the file upon creation.

* **File Type**
> Creates an Excel or CSV file.

### BOM additions

* **Logo**
> You can add your own logo to the BOM (Excel only).

* **Date Created**
> dds BOM creation date.

* **Project name**
> Adds the name of the project based on which the BOM was created.

### Supported physical options

* **Include volume**
> Includes the accumulated volume for all bodies at first level whithin a component.

* **Include area**
> Includes the accumulated area for all bodies at first level whithin a component.

* **Include mass**
> Includes the accumulated mass for all bodies at first level whithin a component.

* **Include density**
> Add's the density of the first body at first level found whithin a component.

* **Include material**
> Includes the material names as an comma seperated list for all bodies at first level whithin a component.

* **Include description**
> Includes the component description. To edit, right click on a component and select _Properties_ in the submenu.



---

<a id="installation"></a>

## Installation

1. Checkout the repository from Github or get the ZIP-package [here](http://www.github.de/macmanpb/BOM-Excel/archive/master.zip)
2. If you have checked out the repo, you can skip point 3
3. Extract the content of the downloaded ZIP to a preferred location
4. Open Fusion360 and load the Add-Ins dialog

	![Toolbar top right](resources/BOM-Excel/toolbar.png)

5. To add the BOM-Excel Add-In, click on the Add-Ins tab and then on the small plus icon.

	![Add-Ins dialog](resources/BOM-Excel/addins_dialog.png)

6. Locate the unzipped _BOM-Excel-master_ folder, open it and choose _BOM-Excel.py_ and click **OK**

7. The Add-In is now listed but not running. Select the _BOM-Excel_ entry, activate _Run on Startup_ and click _Run_

	![Add-In in the list](resources/BOM-Excel/addins-dialog-listed.png)

After _Run_ was clicked the Add-Ins dialog closes automatically.
Check the _Create_ toolbar panel! BOM-Excel is ready to use ;-)

![](resources/BOM-Excel/create_panel.png)




