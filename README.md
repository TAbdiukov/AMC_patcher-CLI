# AppModeChanger patcher - CLI

Patches the app's SUBSYSTEM flag to change its behaviour

## Trivia

There's an amazing Nirsoft VB snippet of AppModeChanger that does not support CLI usage (unless you want to mess with hotkeys, which is not a stable solution). I rewrote the code for command-line usage, but then turns out Nirsoft don't use any Git repos anywhere (from what I could see) I'd have to just host code and stuff here. The name was trunkated to *amc* for the easier CLI usage

Note that if you make changes to the AMC code and compile it, you'd need another copy of AMC, either CLI or GUI, to patch for it to work. Further explained in the **How to compile** section.

## Usage:

	amc <path_to_app> <new_mode>

### Examples:

	amc "C:/Projects/My supa CLI project/Project1.exe" 3 
(to set the Project1 application to the CLI mode), or

	amc "C:/Projects/My crazy CLI project converted to XBOX GUI/app.exe" 14
(to convert the app.exe to XBOX GUI app

## Manual:
    <path_to_app> - Path to your executable. "-tolerable

    <new_mode> - New app SUBSYSTEM mode to set
    
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


## How to compile?

1. *[Recommended for compatibility]* Get a Windows XP VM
2. Get a **Microsoft Visual Basic 6.0** 

***Tip:** I unofficially recommend a portable version sticking around on BT, as you won't have to mess around with the installation and registry. Plus, it's only a few megabytes. Check out **Portable Microsoft Visual Basic 6.0 SP6***

3. Fire up **Microsoft Visual Basic 6.0**, open up the project.
4. Go to File -> Make *.exe -> Save

5. Due to the "[chicken&egg problem](https://en.wikipedia.org/wiki/Bootstrapping_(compilers))" present here, you will have to patch the app with either its other pre-compiled AMC copy or the original NirSoft's GUI [Application Mode Changer](http://www.nirsoft.net/vb/console.zip) ([info](http://www.nirsoft.net/vb/console.html)). By calling:

		amc "path_to_my_new_amc_exe" 3
6. Done!
