# AppModeChanger patcher – CLI

![HMM](icons8-insulin-pen-90.png)
Patches app's SUBSYSTEM flag to modify app's behavior.

[![Download](https://img.shields.io/badge/download-success?style=for-the-badge&logo=github&logoColor=white)](https://github.com/TAbdiukov/AMC_patcher-CLI/releases/download/1.12/amc.exe)


## Trivia

There is an amazing Nirsoft VB snippet of AppModeChanger that does not support CLI usage (unless you want to use hotkeys, which is not a scalable solution). I rewrote the code for command-line usage. The name was truncated to *amc* for the easier CLI usage.

As explained in the **How to compile** section, if you make changes to the AMC code, then after compiling another copy of AMC is needed, either CLI or GUI, to patch it to work.

## Usage:
![Nirsoft original screenshot](https://www.nirsoft.net/vb/console1.gif)


	amc <path_to_app> <new_mode>

### Examples:

	amc "C:/Projects/My CLI project/Project1.exe" 3 
to set the Project1 application to the CLI mode, or

	amc "C:/Projects/My CLI project converted to XBOX GUI/app.exe" 14
to convert the app.exe to XBOX GUI app

## Manual:
    <path_to_app> – Path to your executable. "-tolerable

    <new_mode> – New app SUBSYSTEM mode to set
    
    Informally, one'd need to only know of modes: 2 (CLI) and 3 (GUI)
    
    But below all known modes are listed:
            * 0 UNKNOWN Unknown system
            * 1 NATIVE Image doesn't require a subsystem
            * 2 WINDOWS_GUI Image runs in the Windows GUI subsystem
            * 3 WINDOWS_CUI Image runs in the Windows CLI subsystem (console)
            * 5 OS2_CUI image runs in the OS/2 character subsystem
            * 7 POSIX_CUI image runs in the Posix character subsystem
            * 8 NATIVE_WINDOWS image is a native Win9x driver
            * 9 WINDOWS_CE_GUI Image runs in the Windows CE subsystem
            * 10 EFI_APPLICATION An Extensible Firmware Interface (EFI) application
            * 11 EFI_BOOT_SERVICE_DRIVER An EFI driver with boot services
            * 12 EFI_RUNTIME_DRIVER An EFI driver with run-time services
            * 13 EFI_ROM An EFI ROM image
            * 14 IMAGE_SUBSYSTEM_XBOX No description
            * 16 IMAGE_SUBSYSTEM_WINDOWS_BOOT_APPLICATION No description

## How to compile
1. *[Recommended for compatibility]* Get a Windows XP VM
2. Get **Microsoft Visual Basic 6.0** 

	* **Tip:** There is a portable build, only a few megabytes. Look up <ins>Portable Microsoft Visual Basic 6.0 SP6</ins>

3. Start **Microsoft Visual Basic 6.0**, open the project.
4. Go to File → Make *.exe → Save
5. Patch the app for CLI use:
	* You can use this [AMC patcher](https://github.com/TAbdiukov/AMC_patcher-CLI). For example,

		```
		amc "path_to_my_new_amc_exe" 3
		```
		
	* Or you can use the original Nirsoft's [Application Mode Changer](http://www.nirsoft.net/vb/console.zip) ([docs](http://www.nirsoft.net/vb/console.html)), unpack the archive and then run **appmodechange.exe**

6. Done!


## How to turn your VB6 app into console/CLI
#### (based on my Stackoverflow answer)

1. Clone this repo.
2. Copy `CLI.bas` to your project, then add `CLI.bas` to your project. Now you can use the CLI functions. For example,

```
CLI.setup ' required line, to set up variables
CLI.send "some text -> "
CLI.SetTextColour CLI.FOREGROUND_RED OR CLI.FOREGROUND_GREEN OR CLI.FOREGROUND_INTENSITY ' for Orange and Intensive text, need to use OR, not AND
CLI.sendln "Orange foobar!"
CLI.sendln "maybe another line, why not?"
```

3. Now you can use these functions if your code. *Make sure to call `CLI.setup` first.* First time? You can just test the functionality in `Form1_Load()`
4. Compile your executable via VB6 suite, but the process is not finished yet.
5. Your compiled app has to be patched to work in command-line mode. To do so, `CD` into `AMC_patcher-CLI` folder you called and perform.

        amc "C:/Projects/My CLI project/Project1.exe" 3

Where `"C:/Projects/My CLI project/Project1.exe"` – is the path to your application .exe

Or alternatively for the patching step, use [Nirsoft's original GUI patcher implementation](http://www.nirsoft.net/vb/console.zip). It is less scalable, but just as sturdy.

## Acknowledgements

* Insulin Pen icon by [Icons8](https://icons8.com)
    * Although I have their subscription, better safe than sorry

## See also
*My other small Windows tools,*  

* **<ins>AMC_patcher-CLI</ins>** – (CLI) Patches app's SUBSYSTEM flag to modify app's behavior.
* [exe2wordsize](https://github.com/TAbdiukov/exe2wordsize) – (CLI) Detects Windows-compatible application bitness, without ever running it.
* [SCAPTURE.EXE](https://github.com/TAbdiukov/SCAPTURE.EXE) – (GUI) Simple screen-capturing tool for embedded systems.
* [txtShear](https://github.com/TAbdiukov/txtShear) – (CLI) Fast engine to skew (or shear) text by a desired angle, emulating handwriting.
