# tabswitcher
### Switches Tabs in Microsoft Edge in a fixed time interval 
This is a PowerShell script that continuously switches between the Edge browser and refreshes the active page.

Here's a breakdown of what the script does:

It creates a new instance of the wscript.shell object using the ```New-Object``` cmdlet.
It activates the Edge browser using the ```AppActivate``` method of the ```wscript.shell``` object.
It waits for 5 seconds using the ```Sleep``` cmdlet.
It switches to the next tab in the browser using the ```SendKeys``` method and the Ctrl + TAB keyboard shortcut.
It refreshes the active page using the ```SendKeys``` method and the F5 key.
The script runs in an infinite loop ```(while(1 -eq 1))``` to continuously switch and refresh the active page.
## How to use
Download the script ```tabswitcher.ps1``` and run it
