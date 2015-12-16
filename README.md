# Sams-PowerPoint-Toolbar
Sam's PowerPoint Toolbar. New toolbar group with useful buttons and macros. Instructions and toolbar work only in MS Windows. Tested in Windows 10 but likely also works in earlier versions.

### How to install the toolbar
TODO

### How to create Toolbar file from scratch
1. Start a blank presentation
2. Open Visual Basic editor and import all macros (*.bas) from `/src` directory
3. Save file as PowerPoint Macro-Enabled Presentation (*.pptm)
4. Close PowerPoint and open file in [Office Ribbon Editor](http://www.majorgeeks.com/files/details/office_ribbon_editor.html). For some reason I need to run the Office Ribbon Editor as administrator to get it work
5. Use Document Explorer to *Add 2010 CustomUI*
6. Open RibbonX14, choose *Import CustomUI* and select `/src/RibbonX14.customui`
7. Go to Images tab and *Add images from file system*. Choose all images in `/icons` directory
8. Save and exit Office Ribbon Editor
9. Open saved file and enable macros. You should see the new *Sam's toolkit* group in the toolbar
10. Save the file as PowerPoint Add-In (\*.ppam). Use the *Sam's toolkit.ppam* filename

More detailed instructions on how to add new toobars
* [Stack Overflow: How to add tabs to PowerPoint 2010 that call macros] (http://stackoverflow.com/questions/3867400/how-to-add-tabs-to-powerpoint-2010-that-call-macros/3878978#3878978) 
* [How To Create A PowerPoint 2010 Add-In Using VBA With Custom Buttons In The Ribbon] (http://www.free-power-point-templates.com/articles/create-powerpoint-2010-add-in-with-vba-custom-buttons-ribbon/)

### How to create Installer file from scratch (Install Sam's toolkit.ppsm)
1. Start a blank presentation
2. Import macros and installer picture from `/installer_src` directory
3. Select picture, then Insert -> Action; on Mouse Click tab select *Run macro* and *InstallSamsToolkit*
4. Create text link and do the same as for the picture
5. Save file as PowerPoint Macro-Enabled Show (*.ppsm)
