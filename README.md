# NotARobot
Emulate user activity randomly opening apps through Explorer to generate "legit noise" for EDR and any other log-type collection technology.

## Goal
Be a solid, feature-rich and more natural* alternative to Sheepl, Invoke-UserSimulator and others.\
\* _meaning it uses Explorer navigation instead of launchers and/or powershell etc._\

It probably still need some adjustments to avoid weird parent-child process events. _**This is a working in progress. Check the release notes.**_

## Description
Open Explorer and navigate into binary directory, then hit enters to open it (child process of explorer.exe, as natural as possible).
Then, it waits a random window of time and run another program (or kill one of the previously opened processes).

Considerations: 
1) All the routines that creates and deletes a file uses current user temp folder.
2) All the pseudo random websites (for Edge) are in variables. Feel free to add/change them.
3) I'll try to make some kind of portable library of functions to make it easy to add any application later on. Maybe explaining in the Wiki page so people can fork and contribute with more simulations.

**Current software included (Expand to see a Demo and TODO):**
<details>
  <summary>1) MS Edge</summary>

![Edge Demo](https://github.com/0xleone/NotARobot/blob/main/Edge.gif)
 
</details>
<details>
  <summary>2) MS Office365 (Outlook, Word, Excel)</summary>
  
![Edge Word](https://github.com/0xleone/NotARobot/blob/main/Word.gif)
 
</details>
<details>
  <summary>3) Notepad</summary>

![Notepad Demo](https://github.com/0xleone/NotARobot/blob/main/Notepad.gif)
 
</details>
<details>
  <summary>4) Calc</summary>

![Calc Demo](https://github.com/0xleone/NotARobot/blob/main/Calc.gif)

</details>

## Requirements:
• OS: Win10\
• Win10: Power & Sleep settings: Never/Never\
• MS Edge: Page Layout: Custom: Disable both checkboxes. Background: Disabled; Content: Disabled.\
• MS Office365's Outlook: E-mail pre-configured. Otherwise read and send e-mails will fail.\
• MS Office365's Excel, MS Office365's Word: Nothing specific. It must open without any warnings or prompts and able to edit files (obviously don't use with read-only MS Office version)\
• Notepad, Calc, SnippingTool: Native

## Install
Download compiled version from [here](https://github.com/0xleone/NotARobot/releases) or download AutoIT v3, do your changes and compile yourself into a binary file.

## Usage
Just run the compiled app. It will randomly close after a while (Exit Status = 0).\
If you want to force closing it, you can:
1) **(Recommended)** : Hit the UI close button (Exit Status = 1)
2) **(Experimental)** : Hit *SHIFT+ESC* hotkey (Exit Status = 2)

The hotkey is **NOT** reliable since this project does a LOT of typing emulation and it can fail triggering the hotkey function.

## CURRENT TODO (DEC/2021):
• Add Snipping Tool: Make a screenshot\
• Word/Excel: Create a random file\
• Word/Excel: Delete a random created file\
• Outlook sending random e-mail to itself and/or to a disposable e-mail (e.g. temp-mail.org, 10minutemail.com)\
• Outlook opening random e-mail and attachments
