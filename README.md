# NotARobot
Emulate user activity randomly opening apps through Explorer to generate "legit noise" for EDR etc.

## Install
Download compiled version from HERE or download AutoIT v3, do your changes and compile yourself into a binary file.

## Usage
Just run the compiled app. It will randomly close after a while (Exit Status = 0).
If you want to force closing it, you can hit the UI close button (Exit Status = 1) - more reliable - or hit SHIFT+ESC (Exit Status = 2). The hotkey is NOT reliable since this project does a LOT of typing emulation and it can fail triggering the hotkey.

## Requirements:
• OS: Win10
• Win10: Power & Sleep settings: Never/Never
• Edge: Page Layout: Custom: Disable both checkboxes. Background: Disabled; Content: Disabled.
• MS Office365's Outlook: E-mail pre-configured. Otherwise read and send e-mails will fail.
• MS Office365's Excel, MS Office365's Word: Nothing specific. It must open without any warnings or prompts and able to edit files (obviously don't use with read-only MS Office version)
• Notepad, Calc, SnippingTool: Native
