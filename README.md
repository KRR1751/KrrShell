# KrrShell
Have you wonder, how it will be look like if you can modify your Taskbar, Start menu and Desktop up to you? Instead of downloading some addons like "Open Shell", "Start11" etc.?

Now with this project it is possible!
You can change **ANY LOOK, ANY TEXT as YOU LIKE.**

# How I made this program
Here is the Full playlist of my Livestreams!
https://www.youtube.com/playlist?list=PLfb8PhU_qV0lHSCAV-XUtmqGOzowBT2jD

Or here is the Full playlist how to create yourself one!
https://www.youtube.com/playlist?list=PLfb8PhU_qV0m6-6hrCgv7JG3jrUMjGDrC

# How I install it?
> [!WARNING]
> IF YOU WANT TO USE THE SHELL AS YOUR MAIN ONE, USE IT AT YOUR OWN RISK! IT IS NOT FULLY RELEASED YET AND IT MAY CONTAIN BUGS OR IT CAN EVEN WIPE SOME OF YOUR FILES IF YOU DIDN'T USE IT CORRECTLY.

> [!TIP]
> It is highly recommended before you do this, to make sure you have created System restore.

## To install the Shell yourself, here is a quick tutorial how to install it manually:

> 0. Check if you have installed [.NET Framework 4.8](https://support.microsoft.com/en-us/topic/microsoft-net-framework-4-8-offline-installer-for-windows-9d23f658-3b97-68ab-d013-aa3c3e7495e0) or this Shell will **NOT WORK!**
> 1. Import first a *.reg file called: **"Import_me.reg"** to the Registry
> (This will create a new Subkey "HKEY_CURRENT_USER\Software\Shell" where then all the settings will be stored. *(Planning to have it in the KrrShell already...)*)
> 2. Move (or Copy) folder "NewShell" to your Windows directory (C:\Windows, %WinDir%)
> 3. Press WIN+R and into the Dialog type: "shell:startup", It will open a folder called "Startup"
> 4. Move "ShellActivation.cmd" to the opened Startup folder (ShellActivation.cmd does only the Explorer kill process and the new shell run on each startup of your Windows)
> 5. Now save all your work on all Explorer processes and let only the "Startup" folder opened
> 6. Run the "ShellActivation.cmd" and Done!

> [!NOTE]
> You will not modify anything in your system. When you want to remove it, just simply remove the folder you copied to the %WinDir% and done! Any changes are working in the program it self, not to worry that by this process you'll break something...

# FAQ
**1. Why in Visual Basic?**
- Because I did on it 6 years of experience (from year 2019) and I don't want to learn another programming language...

**2. Can I run it on older Windowses? (such as Windows XP etc.?)**
- This shell runs on .NET Framework 4.8, so any Windows who support this, will also be able to run the shell.

**3. Does it work on Linux??**
- Yes, as I said. If you use something with the .NET Framework 4.8, it will work. So with WineHQ you can even run this shell on Linux!... But you'll not recieve the best experience as when it will be running on Windows.

*Any problem with it or with another question to say, please contact me on: minedowskrr@gmail.com email.*

----------

For supporting the project:
- You can become a [YouTube channel member](https://www.youtube.com/channel/UCBrzUsNl2ZUegwEAkZlRW_Q/join)
- Or you can send me a [Donate here!](https://streamelements.com/krr1751/tip)

Thank you ALL for the support!!!
