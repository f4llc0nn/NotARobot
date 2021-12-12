# NotARobot
Emulate user activity randomly opening apps through Explorer to generate "legit noise" for EDR etc.

## Features
Open Explorer and navigate into binary directory, then hit enters to open it (child process of explorer.exe, as natural as possible).\
**Current software included:**
1) MS Edge
2) MS Office365 (Outlook, Word, Excel)
3) Notepad
4) Calc
5) Snipping Tool


## Install
Download compiled version from HERE or download AutoIT v3, do your changes and compile yourself into a binary file.

## Usage
Just run the compiled app. It will randomly close after a while (Exit Status = 0).\
If you want to force closing it, you can:
1) **(Recommended)** : Hit the UI close button (Exit Status = 1)
2) **(Experimental)** : Hit *SHIFT+ESC* hotkey (Exit Status = 2)

The hotkey is **NOT** reliable since this project does a LOT of typing emulation and it can fail triggering the hotkey function.

## Requirements:
• OS: Win10\
• Win10: Power & Sleep settings: Never/Never\
• MS Edge: Page Layout: Custom: Disable both checkboxes. Background: Disabled; Content: Disabled.\
• MS Office365's Outlook: E-mail pre-configured. Otherwise read and send e-mails will fail.\
• MS Office365's Excel, MS Office365's Word: Nothing specific. It must open without any warnings or prompts and able to edit files (obviously don't use with read-only MS Office version)\
• Notepad, Calc, SnippingTool: Native\

## TODO:
• Snipping Tool: Make a screenshot\
• Word/Excel: Create a random file\
• Word/Excel: Delete a random created file\
• Outlook sending random e-mail to itself and/or to a disposable e-mail (e.g. 10MinuteMail.com)\
• Outlook opening random e-mail and attachments
