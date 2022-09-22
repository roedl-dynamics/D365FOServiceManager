# D365FOServiceManager

Start and stop any windows service direclty via the GUI. Progress of the started/stopped services is shown with a progress bar. Includes a list of added services with their respective status.


# How to use

1.  Run D365FTServiceManager.exe (will ask to be run as administrator).
2.  Once the GUI opened, you can start the predefined services (if you have access to these) by clicking the "Start" button.
3.  You can use the "Stop" button to stop all services included in the list.
4.  You can also add additional services by typing the name of the service in the input and clicking the "Add" button or by pressing <kbd>Enter</kbd>.
5.  You can also select, or multiselect with <kbd>Ctrl</kbd> + <kbd>LMB</kbd>, any service in the list and remove them via the "Remove" button. This just removes them from the list and **does not** stop them.
6.  You can also Start/Stop/Remove services individually with the context menu by right clicking them.

![image](https://user-images.githubusercontent.com/112094138/191783407-28ca1bf7-e66f-4b15-828b-58c313e8eb26.png)
# Edit the script

This program is purely written with the AutoIt script editor SciTE. 
Both are free to download from the following links:
1.  [AutoIt3](https://www.autoitscript.com/site/autoit/downloads/).
2.  [SciTE](https://www.autoitscript.com/site/autoit-script-editor/downloads/).

Once installed, open the D365FOServiceManager.au3 file with SciTE to open and edit the code. By clicking <kbd>F5</kbd>, you can testrun the script.
For further informations have a look at the official AutoIt [website](https://www.autoitscript.com/site/autoit-script-editor/installation/).
