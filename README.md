# WIM-Tools
WIM (Windows Image Management) Tools is a set of tools to create, modify, and manage your Windows images

NOTE: You can download everything needed for this project by simply grabbing the WIM_Tools.zip. Technically, you will only need the WIM Tools.exe but the .zip file will provide the source code and User Guide as well.  

Version: 20.0.1.200  
Release Date: Mar 20, 2022  


*****************************************************************
Here Are Of Some Of The Things That You Can Do With WIM Tools
*****************************************************************

- NEW - For those times where you simply don't have a Windows image with an install.wim file, the program now has a feature to create a new ISO image with an install.wim file by converting the original image with the install.esd file(s) and creating a new ISO image file.

- IMPROVED - The built-in help system has been revised so that help topic menus have the same menu number as the items on the main menu. Previously the help topic numbers were offset by one. While this worked perfectly, it might be less confusing for numbers on the main menu to align perfectly with those on the help menu.

- NEW - The program can now take multiple Windows images as well as Win PE / RE images and create a single media that can be booted from any of those images. That same media can hold additional generic partitions which allow for the storage of other data on the same bootable media.

- Inject Windows updates into one or more editions of Windows (Home, Pro, Education, etc.) and then combine them all into a single Windows image.

- Inject drivers including boot-critical drivers into one or more editions of Windows (Home, Pro, Education, etc.) and then combine them all into a single Windows image.

- Take a Windows image and create a bootable disk from it. This bootable disk will be bootable from BIOS and UEFI based systems, both x64 and x86, it will support files larger than 4 GB, and will boot on systems that don't like to boot from NTFS formatted media. In addition, this tool will allow for the creation of additional partitions on your media that can be optionally BitLocker encrypted. The Windows images saved can be refreshed without affecting the other partitions on the media.

- If you have a Windows image extracted to a disk for modification or to add / remove files, WIM Tools can create a new bootable image from those files.

- WIM Tools can create a new Windows image that pulls editions of Windows from several sources and combines them all into a single image.

- The contents of a Windows image can be re-organized and images can be added or removed from the image.

- Windows image metadata can be altered

- Your images can include both x64 and x86 editions of Windows in the same image.

- Several additional tools are included that will help you manage your WIM Tools projects.

***********************
System Requirements
***********************

WIM Tools will require that the Windows ADK be installed on your system. Only the "Deployment Tools" component needs to be installed. The program will display a warning when it is started if the ADK is not installed. However, it will continue to operate since some functions will work without the ADK. If the user selects a feature from the menu that requires the ADK, the user will be warned and returned to the main menu.

You can download or install the Windows ADK from here:

https://docs.microsoft.com/en-us/windows-hardware/get-started/adk-install

Run the program locally, not from a network location. I have not tested the program or designed it to run from a network location. 

When operating on multiple editions of Windows in the same project (for example, Win 10 Pro, Home, Education editions, etc.), this program is designed to work with editions of the same version. For example, you do not want to mix version 20H2 and 21H1 in the same project. In addition, you should only create ISO images where all Windows editions are of the same build number.

No additional memory beyond the standard requirements for Windows is needed.

We will need a minimum of three times the amount of space your largest project requires. For example, if you are woring a 10 GB Windows image, you will need

- About 10 GB space to extract the image to your hard disk or SSD.
- About 10 GB of temporary storage space to assemble your final project.
- About 10 GB of space to store the final image.
- When working with Windows updates or drivers to inject into your images, you will also need space for these elements.

Note that these items do not all need to be located on the same drive.

Please review the User Guide for additional information and details that will allow you to get the most out of this program.

*********************************
About the Author and the Code
*********************************

I am a Windows enthusiast who started with Windows way back in the '80s when Windows was first introduced. I was a Technical Support Engineer with Microsoft for over ten years. Recently, I've become interested in managing and deploying Windows using nothing but native Microsoft tools. Unfortunately, the management of Windows images can be extremely time consuming because of the sheer number of very lengthy commands that need to be run.

This program began as a series of batch files that allowed me to experiment and slowly build upon each previous step. Eventually, I got tired of the awkward and time consuming experience of programming with batch files. I'm not a professional programmer and the time being wasted trying to work out some of the oddities of batch file programming was taking me away from progressing with other things. It was then that I discovered how simple it would be to script all the commands that I needed to run and at the same time create a far better user experience using QB64, especially since I had played extensivly with Microsoft's Quick BASIC before my Windows days. QB64 is very similar in many ways to the original Quick BASIC but provides the advantages of a modern compiler. Another nice thing is that the entire program resides within a single executable file with no additional support files needed.

Slowly but surely, I kept adding new functionality to the program. The result is the collection of tools all contained within a single executable file that you will find here.

This project is made up of the following files:

WIM Tools.exe - Technically, this is the file needed.  
WIM Tools.bas - The source code. This is a plain text file and can be compiled with the QB64 compiler which can be found at www.qb64.org.  
WIM Tools User Guide.pdf - While help does exist within the program, this user guide may be a handy reference to familiarize you with the program.  

WIM_Tools.zip - Contains all the above files in a single archive.
