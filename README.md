# KrrShell
Have you wonder, how it will be look like if you can modify your Taskbar, Start menu and Desktop completelly up to you? Instead relying to some addons like "Open Shell", "Start11" that just goes to the current explorer etc.?

Now with this project it is possible!
You can change **ANY LOOK, ANY TEXT as YOU LIKE.**

## How it was made?
Here is a [full playlist](https://www.youtube.com/playlist?list=PLfb8PhU_qV0lHSCAV-XUtmqGOzowBT2jD) of my Livestreams where I made all the work!

Or here is a [full playlist](https://www.youtube.com/playlist?list=PLfb8PhU_qV0m6-6hrCgv7JG3jrUMjGDrC) how to create one yourself!

## How do I install it?
> [!WARNING]
> Please follow those steps carefully because it may not be working at all and it is mainly made for advanced users at this point. Well the "setup part" as the Shell installer included there may not be working so it can happen you need to install it manually.
> 
> It is also highly recommended before you do this, to make a System restore point if anything will happen.

> [!IMPORTANT]
> Check if you have installed [.NET Framework 4.8](https://support.microsoft.com/en-us/topic/microsoft-net-framework-4-8-offline-installer-for-windows-9d23f658-3b97-68ab-d013-aa3c3e7495e0) or this Shell will **NOT WORK!**

<sub>This is a tutorial how to install it manually if the Shell Installer will not work:</sub>

| Numbers | Action |
| --- | --- |
| 1. | Import first a *.reg file called: **"Import_me.reg"** to the Registry *(This will create a new Subkey* `HKEY_CURRENT_USER\Software\Shell` *where then all the settings will be stored)* |
| 2. | Move (or Copy) folder "NewShell" to your Windows directory `C:\Windows, %WinDir%` |
| 3. | Press WIN+R and into the Dialog type `shell:startup`, It will open a folder called "Startup" |
| 4. | Move `ShellActivation.cmd` to the opened Startup folder *(ShellActivation.cmd does only the Explorer kill process and the new shell run on each startup of your Windows)* |
| 5. | Now save all your work on all Explorer processes and let only the `Startup` folder opened |
| 6. | Run the `ShellActivation.cmd` and Done! |

> [!NOTE]
> By this process you'll not modify anything in your system. For a removal, you'll only remove the folder you moved/copied to the `C:\Windows, %WinDir%` (the `NewShell`) and **done!**
> 
> Any changes, setup etc. are working in the program itself. Not to worry that by this process you'll break something...

# FAQ
**1. Why in Visual Basic.net?**
- Because my experience (from year 2019) with all the programming learning, understanding the Win32APIs are coming from this language as I know the most of it (even without AI)

**2. Can I run it on older Windowses? (such as Windows XP etc.?)**
- This shell runs on .NET Framework 4.8, so any Windows who support this, will also be able to run the shell.

**3. Does it work on Linux??**
- Yes, as I said. If you use something with the .NET Framework 4.8, it will work. So with WineHQ you can even run this shell on Linux!... But you'll not recieve the best experience as when it will be running on Windows.

**4. Did you used AI?**
- Yes, but only as a "tool" not just saying to it "generate me a shell" and call it a day. Those features inside were my exact ideas and with AI it only made the process faster. I could make it only just myself, but it would take me a whole decade. AI is good if you use it the "good" way and still use "your ideas" instead relying anything up to it.


## 💸 Support this project!
- You can become a [YouTube channel member](https://www.youtube.com/channel/UCBrzUsNl2ZUegwEAkZlRW_Q/join)
- Or for a mini-tip, [Donate here!](https://streamelements.com/krr1751/tip)



*Any problem with it or with another question to say, please contact me on: minedowskrr@gmail.com email.*

Thank you ALL for the support!!!
