# Sams-PowerPoint-Toolbar
Sam's PowerPoint Toolbar. New toolbar group with useful buttons and macros. Instructions and toolbar work only in MS Windows. Tested in Windows 10 but likely also works in earlier versions.

### How to create Toolbar file from scratch
1. Start a blank presentation
2. Import macros from /macros directory
3. Save file as PowerPoint Macro-Enabled Presentation (*.pptm)
4. Close PowerPoint and open file in [Office Ribbon Editor](http://www.majorgeeks.com/files/details/office_ribbon_editor.html). For some reason I need to run the Office Ribbon Editor as administrator to get it work
5. Using Document Explorer "Add 2010 CustomUI"
6. Open Ribbon14, choose Import CustomUI and select /macros/RibbonX14.customui
7. Go to Images tab and Add images from file system. Choose all images in /icons directory
8. Save and exit Office Ribbon Editor
9. Open saved file and enable macros. You should see the new *Sam's toolkit* group in the toolbar
10. Save the file as PowerPoint Add-In (*.ppam). Use the *Sam's toolkit.ppam* filename

### How to create Installer from scratch (Install Sam's toolkit.ppsm)
1. Start a blank presentation
2. Import macros and installer picture from /installer_src directory
3. Select picture, then Insert -> Action; on Mouse Click tab select Run macro and InstallSamsToolkit
4. Create text link and do the same as for the picture
5. Save file as PowerPoint Macro-Enabled Show (*.ppsm)
