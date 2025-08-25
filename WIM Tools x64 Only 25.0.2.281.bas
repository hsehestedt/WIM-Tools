' USE QB64PE TO COMPILE
' https://github.com/QB64-Phoenix-Edition/QB64pe/releases


' WIM (Windows Image Manager) Tools, x64 Only Edition
' (c) 2025 by Hannes Sehestedt

' Release notes can be found at the very end of the program

' This program is intended to be run on Windows x64 and should be compiled with the 64-bit version of QB64PE.


Option _Explicit

' ********************************************************
' ** Make sure to keep the "$VersionInfo" updated below **
' ********************************************************

' Perform some initial setup for the program

' This program makes use of "wimlib-imagex.exe" and "libwim-15.dll" to modify the install.wim
' so that the product type is set to "server". Doing this will allow an upgrade install of
' Windows on unsupported hardware. We are embedding those files into this program and we will
' extract them if and when they are needed.

Rem $Dynamic
$ExeIcon:'WIM Tools.ico'
$Embed:'.\Resources\wimlib-imagex.exe','WimLib'
$Embed:'.\Resources\libwim-15.dll','libwim15'
$VersionInfo:CompanyName='Hannes Sehestedt'
$VersionInfo:FILEVERSION#=25,0,2,281
$VersionInfo:ProductName='WIM Tools x64 Only Edition'
$VersionInfo:LegalCopyright='(c) 2025 by Hannes Sehestedt'
$Console:Only
_Source _Console
Width 120, 30
Option Base 1

BeginProgram:

' IMPORTANT: When you return here with a "GOTO BeginProgram" ALWAYS preceed this with "CHDIR ProgramStartDir$" like this:
'
' CHDIR ProgramStartDir$: GOTO BeginProgram
'
' This ensures that we change the current working directory back to the original location before we jump back here and
' clear the variable that kept track of that location.

Cls
Color 15
Print "Initializing program. Standby..."

' Throughout the program, when we finish running a routine, rather than simply returning to the main menu, we will
' return to this point. By doing so, we will clear all variables and redimension everything from scratch. Even
' though this program is no resource hog, this will keep things tiday and clean, especially if the user runs
' multiple different routines in succession.

Clear

Dim Cmd As String ' Cmd$ will hold commands that we will run via a shell command
Dim Temp1 As String ' Temp1$ is for holding temporary data. Don't use long term because other parts of program will use it
Dim Temp2 As String ' Same as Temp1$
Dim x As Integer ' Generic variable reused throughout program, mainly as a counter in FOR ... NEXT loops

' ***********************************************************
' ***********************************************************
' ** The following strings hold the program version number **
' **               and program release date.               **
' **            Make sure to keep this updated.            **
' ***********************************************************
' ***********************************************************

Dim Shared ProgramVersion As String ' Holds current program version and is displayed in the console title the program
Dim ProgramReleaseDate As String

ProgramVersion$ = "25.0.2.281"
ProgramReleaseDate$ = "Aug 25, 2025"


' ******************************************************************************************************************
' * Color Scheme                                                                                                   *
' *                                                                                                                *
' * COLOR 15   : Normal Text : Color 15: White letters on black background                                         *
' * COLOR 10   : Used to draw user attention, mainly uses for progress status : Green on black background          *
' * COLOR 4    : Used mainly in error messages and some summary status info : Red on black background              *
' * COLOR 0, 9 : Used in Menu : Color 0, 9: Black text on light blue background                                    *
' * COLOR 0, 6 : Used in Menu : Black text on bownish-orange background                                            *
' * COLOR 0, 14: Used in highlighting, mainly for output from command line utils : Black text on yellow background *
' * COLOR 0, 10: Used to draw user attention to certain info : Black text on green background                      *
' * COLOR 14, 4: Used in errors and warnings : Yellow text on red background                                       *
' * COLOR 0, 13: Used for help menus : Black text on purple background                                             *
' ******************************************************************************************************************


' **************************
' * Declare variables here *
' **************************

Dim AddAnswerFile As String
Dim AdditionalPartitions As Integer ' The number of partitions that a user wants to add to a bootable thumb drive
Dim AddPart As String ' Set to either "Y" or "N" depending upon whether user wishes to add additional partitions to a bootable thumb drive
Dim ADKFound As Integer ' This flag will be "0" if the ADK is not found on the system, or "1" if it is found.
Dim ADKLocation As String ' This will hold the ADK path to the Windows ADK
Dim AllFilesAreSameArc As Integer ' Flag is set to "1" if all images in the project are of a single architecture type
Dim Another As String ' Temp variable
Dim AnswerFilePath As String ' Full path to answer file save location for answer file generator. Includes file name
Dim AnswerFilePresent As String
Dim Arc As String ' Used to store architecture type in the routine to create a VHD
Dim Architecture As Integer ' Flag that gets set to 1 for a single architecture image, 2 for dual architecture, and 0 if an invalid image
Dim ArchitectureChoice As String ' Search code for ArchitectureChoice$ for a comment explaining usage
Dim AvailableSpace As Long ' Used for tracking available space on a disk
Dim AvailableSpaceString As String ' Used for tracking available space on a disk
Dim BitLockerCount As Integer ' Stores the number of partitions that need to be encrypted
Dim BootWimMods As String 'Holds a value that determines what mods are made to a Boot.wim
Dim BypassDeviceEncryption As String ' Y or N to indicate whether auto device encryption should be bypassed
Dim BypassQualityUpdatesDuringOobe As String ' Y or N to indicate if installation of quality updates during Windows setup should be bypassed
Dim BypassWinRequirements As String ' Y or N to indicate whther we will bypass Win 11 system requirements
Dim CDROM As String ' The drive letter assigned to the mounted ISO image
Dim ChosenIndex As Integer
Dim Shared CleanupSuccess As Integer ' Set to 0 if we do not successfully clean contents of folder in the cleanup routine, otherwise set to 1 if we succeed
Dim Column As Integer ' Used for positioning cursor on screen
Dim ComputerName As String ' Computer name to be assigned during setup from answer file
Dim Shared CreateEiCfg As String ' Set to "Y" or "N" to indicate whether an ei.cfg file should be created.
Dim CurrentImage As Integer ' A counter used to keep track of the image number being processed
Dim CurrentIndex As String ' A counter used to keep track of the index number within an image being processed
Dim CurrentIndexCount As Integer
Dim CurrentTime As String ' Date and Time combined with a comma as a separator
Dim Description As String ' Holds the description metadata to be assigned to a Windows edition within an image
Dim DescriptionFromFile As String ' Holds the DESCRIPTION field of an image parsed from Image_Info.txt file
Dim DestArcFlag As String ' A flag that varies wit the architecture type used to build out a final path
Dim Destination As String ' Destination path
Dim DestinationFileName As String ' The file name of the ISO image to be created without a path
Dim DestinationFolder As String ' The destination folder where all the folders created by the project will be located as well as the final updated ISO images
Dim DestinationIsRemovable As Integer ' Flag to indicate if the originally specified destination is removable
Dim DestinationPath As String ' The destination path for the ISO image without a file name
Dim DestinationPathAndFile As String ' The full path including the file name of the ISO image to be created
Dim Shared DiskDetail(0) As String ' Stores details about each disk in the system
Dim DiskIdTarget As String ' Disk number to which Windows will be installed with generated answer file
Dim DiskID As Integer ' Used in multiple places to ask the user for a DiskID as presented by the Microsoft DiskPart utility
Dim Shared DiskIDList(0) As Integer ' Used to store a list of valid Disk ID numbers
Dim DiskIDSearchString As String ' Holds a disk ID that will be searched for in the output of diskpart commands
Dim Shared DISM_Error_Found As String ' Holds a "Y" if an error is found in log file, a "N" if not found.
Dim Shared DISMLocation As String ' Holds the location of DISM.EXE as reported by the registry
Dim DisplayName As String ' Full name or Display Name that is shown on lock screen etc. as full name of the user name
Dim DisplayUnit As String ' Holds "MB", "GB", or "TB" to indicate what units user is entering partition size in
Dim DriveLetter As String ' Take a path and store the drive letter from that path (C:, D:, etc.) in this valiable to be used to determine if drive is removable or not
Dim DST As String ' A path that includes location of install.wim files
Dim EditionName As String
Dim EfiParSize As String ' Size of EFI partition in MB
Dim Shared ExcludeAutounattend As String ' If set to "Y" then exclude any existing autounattend.xml file, if set to "N" then it is okay to copy the file
Dim exFATorNTFSdriveletter As String
Dim ExportFolder As String ' Used by the routine for exporting drivers from a system as well as the Reorg routine.
Dim FAT32DriveLetter As String ' Letter assigned to 1st partition
Dim ff As Long ' Holds the value returned by FREEFILE to determine an available file number
Dim ff2 As Long
Dim Shared FileCount As Integer ' The number of ISO image files that need to be processed. In multiboot image creation program, this hold the number of images we have to process.
Dim FileLength As Single
Dim FileName As String ' Used by routine to inject Win 11 requirements bypass registry entries to hold ISO image name to be modified.
Dim FileSourceType As String
Dim FinalImageName As String
Dim FirstLogonCommandCounter As Integer ' Counter
Dim FSType As String ' Set to either NTFS of EXFAT to determine what filesystem user wants to use
Dim Highest_Single As Integer
Dim IDX As String
Dim Shared ImageArchitecture As String ' Used by the DetermineArchitecture routine to determine if an ISO image is x86, x64, or dual architecture
Dim ImageInfo As String
Dim Shared IMAGEXLocation As String ' The location of the ImageX ADK utility as reported by the registry
Dim ImagePath As String
Dim ImageSourceDrive As String
Dim ImageType As String ' For the routine to convert between an install.esd and install.wim, tracks whether image in ISO is an ESD or a WIM file.
Dim Index As String ' Holds index number for the image being processed as a string without leading space
Dim IndexCount As Integer ' Holds the number of indices found in an ISO image
Dim IndexCountString As String ' A string version of IndexCount with leading space stripped
Dim IndexCountLoop As Integer
Dim IndexOrder As String
Dim IndexRange As String ' A temporary string for a user to specify a range of numbers. Example: 1-3 5 7-8
Dim IndexString As String ' This is the value of the integer variable Index converted to a string
Dim IndexVal As Integer
Dim InjectionMode As String ' From the main menu, set to "UPDATES" if user wants to inject Windows updates, or "DRIVERS" if user wants to inject drivers.
Dim InstallFile As String
Dim InstallFileTest As String
Dim InstallPar As String ' The partition number to which Windows should be installed for answer file
Dim Shared IsRemovable As Integer ' Value returned subroutine to determine if a disk ID or drive letter passed to it is removable or not
Dim LCU_Update_Avail As String
Dim LettersAssigned As Integer
Dim LimitWinParSize As String ' Y or N to indicate whether Windows partition size should be limited rather than using all remaining space
Dim Shared ListOfDisks As String
Dim MainLoopCount As Integer ' Counter to indicate which loop we are in.
Dim MakeBootablePath As String
Dim MakeBootableSourceISO As String ' The full path and file name of the ISO image that the user want to make a bootable thumb drive from
Dim ManualAssignment As String
Dim MaxLabelLength As Integer ' The allowable length for a volume label - 11 for exFAT, 32 for NTFS
Dim MediaLetter As String
Dim MenuSelection As Integer ' Will hold the number of the menu option selected by the user
Dim Midnight As Integer ' Just before creating an ISO image, set to "1" if we are within the two seconds prior to midnight, otherwise, set to "0"
Dim MoreFolders As String
Dim Shared MountedImageCDROMID As String ' The MountISO returns this value
Dim Shared MountedImageDriveLetter As String ' The MountISO returns this value
Dim MountDir As String ' Used to hold text while reading from a file looking for a DISM mount location
Dim MsrParSize As String ' Size of MSR partition in MB
Dim Multiplier As Single
Dim NameFromFile As String ' Holds the NAME field of an image parsed from Image_Info.txt file
Dim NewLabel As String
Dim Shared NumberOfDisks ' Stores the number of disk drives that diskpart sees in the system
Dim Shared NumberOfSingleIndices As Integer
Dim Shared NumberOfFiles As Integer ' Used by the FileTypeSearch subroutine to keep count of the number of files found in a folder of the type specified by a user
Dim Shared NumberOfx64Indices As Integer
Dim Shared NumberOfx86Indices As Integer
Dim NumberOfx64Updates As Integer
Dim Shared OpsPending As String
Dim OpsPendingFileCheck As String
Dim OriginalImageType As String ' Hold type of image selected by user
Dim Shared OSCDIMGLocation As String ' Holds the location of OSCDIMG.EXE as reported by the registry
Dim Other_Updates_Avail As String
Dim OutputFileName As String ' For Windows multiboot image program, holds the final name of the ISO image to be created (file name and extension only, no path)
Dim Override As String
Dim ParSizeInMB As String ' Holds the size of a partition as a string
Dim Par1MultiInstancesFound As Integer ' This and next 3 vars get a count of how many single ir multi image partitions exist in system
Dim Par2MultiInstancesFound As Integer
Dim Par1SingleInstancesFound As Integer
Dim Par2SingleInstancesFound As Integer
Dim PE_Files_Avail As String
Dim Phase1Commandcounter As Integer ' Used as a counter
Dim Phase4Commandcounter As Integer ' Used as a counter
Dim PreviousSetup As String ' This will be "Y" if user wants to force use of previous setup for Windows 11, otherwise set to "N"
Dim ProductKey As String ' Generic installtion product key to be used for installation
Dim Shared ProgramStartDir As String ' Holds the original starting location of the program
Dim ProjectArchitecture As String ' In Multiboot program, hold the overall project architecture type (x86, x64, or DUAL)
ReDim Shared RangeArray(0) As Integer ' Each individual numeric value from the range of numbers passed into the ProcessRangeOfNums routine expanded into individual numbers
Dim ReadLine As String
Dim ReorgFileName As String
Dim ReorgSourcePath As String
Dim Resource As String
Dim RowEnd As Integer
Dim Row As Integer ' Used for positioning cursor on screen
Dim SafeOS_DU_Avail As String
Dim SASSU_Update_Avail As String ' This variasble is set to "Y" if a Standalone SSU update is available
Dim Shared ScriptingChoice As String ' Used to track what scripting operation user wishes to perform
Dim Shared ScriptContents As String ' Holds the entire contents of previously created script for playback
Dim Shared ScriptFile As String ' Used to store the name of the script file to be run
Dim SearchPosition As Integer ' Hold the position in which a string was found within a string variable
Dim Setup_DU As String ' Holds the location of the Setup Dynamic update file
Dim Silent As String
Dim Shared ShutdownStatus As Integer
Dim SingleImageTag As String
Dim SingleImageCount As Integer
Dim SingleOrMulti As String ' Set to "SINGLE" to create a single image boot disk or "MULTI" to create a multi image disk
Dim Shared Skip_PE_Updates As String ' if Set to "Y" we will not apply SSU and LCU updates to the WinPE (boot.wim) image
Dim Source As String ' Source file in an ESD to WIM conversion
Dim SourceArcFlag As String
Dim SourceFolder As String ' Will hold a folder name
Dim SourceFolderIsAFile As String ' If SourceFolder$ actually contains a filename rather than a path, set this to "Y", else set to "N"
Dim SourceImage As String
Dim SourcePath As String ' Holds the path containing the files to be injected into an ISO image file
Dim SRC As String
Dim SystemType As String ' Used by Answer File Gen to hold UEFI or BIOS as type of system for which answer file is being made
Dim Shared Temp As String ' Temporary string value that can be shared with subroutines and also used elsewhere as temporary storage
Dim Shared TempArray(100) As String ' Used by FileTypeSearch subroutine to keep the name of each file of type specified by user. We assume that we will need less than 100.
Dim Shared TempLocation As String ' This variable will hold the location of the TEMP directory.
Dim TempLong As Long
Dim TempPath As String ' A temporary variable used while manipulating strings
Dim TempValue As Double
Dim TimeZone As String ' Time zone to be used for answer file generator
Dim TotalImagesToUpdate As Integer
Dim TotalIndexCount As Integer
Dim Shared TotalNumsInArray As Integer
Dim TotalPartitions As Integer ' The total number of partitions that need to be created on a bootable thumb drive
Dim TotalSpaceNeeded As Long
Dim TempPartitionSize As String
Dim TempUnit As String
Dim TotalFiles As Integer
Dim Units As String
Dim UpdateAll As String ' User is asked by routine to inject updates or drivers if all images should be updated. This string hold their response
Dim UpdatesLocation As String
Dim UpdateThisFile As String
Dim UserCanPickFS As String
Dim UserLocale As String ' Holds either en-US or en-001 for answer file generator
Dim UserName As String ' Holds user name for answer file generation
Dim UserSelectedImageName As String ' If the user wants a specific name for the final ISO image, this string will hold that name
Dim ValidDisk As Integer ' Set to 0 if user chooses an invalid Disk ID, 1 if their choice is valid
Dim Shared ValidRange As Integer ' A flag that indicates whether a range of numbers supplied by a user is valid or not. 0=Invalid, 1=Valid
Dim VHDFilename As String
Dim VHDSize As Long
Dim VHDXPath As String ' The path to where a virtual disk drive is to be created
Dim VHDXFileName As String ' The file name to give a virtual disk drive that is to be created
Dim VHDXSize As Long ' The size in MB to create the virtual disk drive
Dim VHDXSizeString As String
Dim VHDXLetter As String ' The drive letter to assign to the virtual disk drive after it is created and mounted
Dim VHD_Type As Integer
Dim VolLabel As String
Dim VolumeLabel As String
Dim VolumeName As String ' Used by the routine to create a generic ISO image to store the volume name that the user would like to assign to the image
Dim WimInfo As String ' Holds lines of text as they are being read from WinInfo.txt file
Dim WimInfoFound As Integer ' A flag used to indicate whether an index specified by user was found successfully in Image_Info.txt file
Dim WinParSize As String ' Size of Windows partition in MB
Dim WinPEFound As Integer ' Set to 1 if Windows PE is installed, 0 if not installed
Dim WinPELocation As String ' Holds the path to the WinPE installation location
Dim WinPE_Temp As String
Dim WinReParSize As String ' Used by Answer File Generator to hold size of WinRE partition in MB
Dim WinREPartitionSize As Long ' Size to make the WinRE partition when creating a VHD and deploying Windows to it
Dim WipeOrRefresh As Integer
Dim x64ExportCount As Integer
Dim x64Updates As String
Dim y As Integer ' General purpose loop counter
Dim Shared YN As String ' This variable is returned by the "YesOrNo" procedure to parse user response to a yes or no prompt. See the SUB procedure for details
Dim z As Integer ' General purpose loop counter

' The following arrays are dimensioned dynamically within the program. Not all of the variables dimensioned within the body of the program may be listed here.

' ArchitectureArray$()
' AutoUnlock$() -  Flag is set to either "Y" or "N" to indicate whether this partition should be autounlocked on current system
' BitLockerFlag$() - Flag is set to either "Y" or "N" to indicate whether this partition should be encrypted
' EditionDescriptionArray$()
' EditionNameArray$()
' FileArray$() - Will hold the name of each ISO image file to be updated.
' FinalFileNameOnly$() - This is the destination file name without a path
' FinalFilePathAndName$() - This is the full path and file name of the destination file to be created
' IndexArray$() - Holds the selected the index number for the Windows edition chosen for each ISO image.
' IndexStringArray$()
' ISOCounterArray()
' MatchArray() - for Multiboot routine, set to 0 if no previous image in the array of images to be processed is both same architecture and from same ISO image,
'                otherwise has a non-zero value (it gets incremented)
' PartitionSize$() - The size to create a partition converted to a string so that leading space is stripped out.
' SingleArray()
' SourceFileNameOnly$() - The name of the source file without a path
' SourcePathArray$()
' WinRE_x64_Present()
' x64Array()

' ********************************
' * End of variable declarations *
' ********************************


_ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " by Hannes Sehestedt"

' Get the location of the TEMP directory and the location where the program was run from.

TempLocation$ = Environ$("TEMP")
ProgramStartDir$ = _CWD$

' Change the current working directory to the TEMP directory. This will cause all the temporary files
' that we write to be created there.

ChDir TempLocation$

' This program needs to be run elevated. Check to see if the program is running elevated, if not,
' restart in elevated mode and terminate current non-elevated program.

' We need to begin by parsing the original command line for a "'" (single quote) character.
' If this character exists, it will cause the command to self-elevate the program to fail.
' We resolve this by changing the single quote character into two single quotes back-to-back.
' This will "escape" the single quote character in the command.

Cls
Print "Verifying that the program is being run elevated."
Print "Please standby..."

Temp1$ = Command$(0)
For x = 1 To Len(Temp1$)
    If Mid$(Temp1$, x, 1) = "'" Then
        Temp2$ = Temp2$ + "''"
    Else
        Temp2$ = Temp2$ + Mid$(Temp1$, x, 1)
    End If
Next x

If (_ShellHide(">nul 2>&1 " + Chr$(34) + "%SYSTEMROOT%\system32\cacls.exe" + Chr$(34) + " " + Chr$(34) + "%SYSTEMROOT%\system32\config\system" + Chr$(34))) <> 0 Then
    ChDir ProgramStartDir$
    Cmd$ = "powershell.exe " + Chr$(34) + "Start-Process '" + (Mid$(Temp2$, _InStrRev(Temp2$, "\") + 1)) + "' -Verb runAs" + Chr$(34)
    Shell Cmd$
    System
End If

' If we reach this point then the program was run elevated.

' Determine the location of the ADK utilities DISM and OSCDIMG.

' We are performing a registry query and redirecting the otput to ADKSearch.txt. If the registry key does not exist an error will be generated
' by the "reg query" command so we have to redirect that to the file by using "2>&1" at the end of the command.
'
' NOTE: We store the location for DISM.EXE in DISMLocation$, the location for OSCDIMG.EXE in OSCDIMGLocation$, IMAGEX.EXE in IMAGEXLocation$.
' We also set a flag called ADKFound to 1 if DISM.EXE is found to indicate that the ADK is installed. We only need to do this after the verification that DISM.EXE is
' installed; no need to repeat that after getting the location for OSCDIMG.EXE.

Cmd$ = "reg query " + Chr$(34) + "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows Kits\Installed Roots" + Chr$(34) + " /v KitsRoot10 > ADKSearch.txt 2>&1"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
Open "ADKSearch.txt" For Input As #1

' Init variables

ADKLocation$ = ""
ADKFound = 0

Do
    Line Input #1, DISMLocation$
    If InStr(DISMLocation$, "REG_SZ") Then
        ADKLocation$ = Right$(DISMLocation$, ((Len(DISMLocation$) - (InStr(DISMLocation$, "REG_SZ") + 9))))
        ADKFound = 1
        Exit Do
    End If
Loop Until EOF(1)

Close #1
Kill "ADKSearch.txt"

If ADKFound = 1 Then
    DISMLocation$ = ADKLocation$ + "Assessment and Deployment Kit\Deployment Tools\amd64\DISM\DISM.exe"
    OSCDIMGLocation$ = ADKLocation$ + "Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\Oscdimg.exe"
    IMAGEXLocation$ = ADKLocation$ + "Assessment and Deployment Kit\Deployment Tools\amd64\DISM\ImageX.exe"
    WinPELocation$ = ADKLocation$ + "Assessment and Deployment Kit\Windows Preinstallation Environment"
    If _DirExists(WinPELocation$) Then
        WinPEFound = 1
    Else
        WinPEFound = 0
    End If
End If

If ADKFound = 0 Then
    Cls
    Color 14, 4: Print "WARNING!": Color 15
    Print
    Print "We did not find the Windows ADK installed on your system. This program requires that the ADK be installed. Note that"
    Print "you only need to install the Deployment Tools option. Everything else can be omitted. There are some features of"
    Print "this program that will work without the ADK, so we will now take you to the main menu."
    Print
    Print "If you try to run a feature that requires the ADK, we will warn you and then take you back to the menu."
    Pause
End If

' Set an AV exclusion to the current program name

Cmd$ = "powershell.exe -command Add-MpPreference -ExclusionProcess " + "'" + Chr$(34) + Command$(0) + Chr$(34) + "'"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

' In the event that the program was previously terminated unexpectedly, the AV exclusion created later in the program
' for the destination folder may not have been cleared. In that case, a temp file that we create a file called
' "WIM_Exclude_Path.txt" would still be present. That file holds the destination location for which we created an
' exclusion. Clear that exclusion now. By doing this, all the user needs to do is to run this program again and it
' will automatically clear the exclustion. Note that this only applies to the routines that inject updates, drivers,
' and boot-critical drivers.

If _FileExists("WIM_Exclude_Path.txt") Then
    ff = FreeFile
    Open "WIM_Exclude_Path.txt" For Input As #ff
    Line Input #ff, Temp$
    Close #ff
    Cmd$ = "powershell.exe -command Remove-MpPreference -ExclusionPath " + "'" + Chr$(34) + Temp$ + Chr$(34) + "'"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Kill "WIM_Exclude_Path.txt"
End If

' If a temporary file named "WIM_File_Copy_Error.txt" still exists from a pevious run of the pogram, delete it.
' This makes no functional difference to the program, it simply cleans up a file that is not needed.

If _FileExists("WIM_File_Copy_Error.txt") Then Kill "WIM_File_Copy_Error.txt"

' If a file by the name of "WIM_Shutdown_log.txt" exists, this means that on the last run of the program
' the user chose to perform a shutdown after the program finished. Before the shutdown the program saves
' any status messages to this file. We will now display that information to the user and then delete the
' file.

If _FileExists("WIM_Shutdown_log.txt") Then
    Cls
    Print "We have detected that the last time this program was run, it was requested that the system be shutdown or"
    Print "hibernated after the program completed. This message indicates that the program ran to completion."
    Color 10
    Print "_________________________________________________________________________________________________________________"
    Color 15
    Print
    Shell "Type WIM_Shutdown_log.txt"
    Print
    Color 10
    Print "_________________________________________________________________________________________________________________"
    Color 15
    Pause
    Kill "WIM_Shutdown_log.txt"
End If

' Display the main menu

MainMenu:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 15
Print
Print
Color 0, 14
Print "    1) Inject Windows updates into one or more Windows editions and create a multi edition bootable image       "
Print "    2) Inject drivers into one or more Windows editions and create a multi edition bootable image               "
Print "    3) Inject boot-critical drivers into one or more Windows editions and create a multi edition bootable image "
Print "    4) Modify Windows ISO to bypass system requirements and optionally force use of previous version of setup   "
Color 0, 10
Print "    5) Make or update a bootable drive from one or more Windows / WinPE / WinRE ISO images                      "; 0
Print "    6) Create a bootable Windows ISO image that can include multiple editions                                   "
Print "    7) Create a bootable ISO image from Windows files in a folder                                               "
Print "    8) Reorganize the contents of a Windows ISO image                                                           "
Print "    9) Convert between an ESD and WIM either standalone or in an ISO image                                      "
Color 0, 3
Print "   10) Get image info - display basic info for each edition in an ISO image and display Windows build number    "
Print "   11) Modify the NAME and DESCRIPTION values for entries in a WIM file                                         "
Color 0, 6
Print "   12) Export drivers from this system                                                                          "
Print "   13) Expand drivers supplied in a .CAB file                                                                   "
Print "   14) Create a Virtual Disk (VHDX) - NOTE: Win 11 23H2+ has a new GUI to make doing this from the OS easy      "
Print "   15) Create a VHD, deploy Windows to it, and add it to the boot menu to make a dual boot configuration        "
Print "   16) Create a generic ISO image and inject files and folders into it                                          "
Print "   17) Cleanup files and folders                                                                                "
Color 15
Color 0, 55
Print "   18) Unattended answer file generator                                                                         "
Color 15
Color 0, 13
Print "   19) Program help                                                                                             "
Color 0, 8
Print "   20) Exit                                                                                                     "
Locate 3, 40: Color 0, 14:
Print "   ";
Color 15
Print " Update Windows Image Tools    ";
Color 0, 10
Print "   ";
Color 15
Print " Bootable Media and Image Creation Tools"
Color 15
Locate 4, 40: Color 0, 3
Print "   ";
Color 15
Print " WIM Metadata Tools            ";
Color 0, 6
Print "   ";
Color 15
Print " Various Helpful Utilities"

Locate 5, 40: Color 0, 55
Print "   ";
Color 15
Print " Answer File Generator         ";

Color 0, 13
Print "   ";
Color 15
Print " Help"
Locate 29, 0
Input "   Please make a selection by number (20 Exits from the program): ", MenuSelection

' Some routines require that the Windows ADK be installed. We will now check to see if the option selected by the user is one of those routines.
' If it is, then we warn the user and return them to the main menu. If the ADK was found, then we skip this check.

If ADKFound = 1 Then GoTo Skip_ADK_Check


Select Case MenuSelection
    Case 1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 15, 16, 17
        Cls
        Color 14, 4: Print "WARNING!";: Color 15: Print " The routine selected from the menu requires the Windows ADK to be installed. We have detected that the ADK"
        Print "is not installed on this system. Please install the ADK and restart this program to perform this operation."
        Pause
        GoTo MainMenu
End Select

Skip_ADK_Check:

Select Case MenuSelection
    Case 1
        InjectionMode$ = "UPDATES"
        ExcludeAutounattend$ = "Y"
        Scripting ("Inject Updates Into Windows")

        ' If the user has elected to play a script, then open the script file now and play it back

        If ScriptingChoice$ = "P" Then
            ScriptContents$ = _ReadFile$("WIM_SCRIPT.TXT")
            Kill ("WIM_SCRIPT.TXT")
            _ScreenPrint ScriptContents$
        End If
        EiCfgHandling
        Skip_PE_Updates_Check
        GoTo InjectUpdates
    Case 2
        InjectionMode$ = "DRIVERS"
        ExcludeAutounattend$ = "Y"
        Scripting ("Inject Drivers Into Windows")

        ' If the user has elected to play a script, then open the script file now and play it back

        If ScriptingChoice$ = "P" Then
            ScriptContents$ = _ReadFile$("WIM_SCRIPT.TXT")
            Kill ("WIM_SCRIPT.TXT")
            _ScreenPrint ScriptContents$
        End If
        EiCfgHandling
        GoTo InjectUpdates
    Case 3
        InjectionMode$ = "BCD"
        ExcludeAutounattend$ = "Y"

        Scripting ("Inject Boot-critical Drivers")

        ' If the user has elected to play a script, then open the script file now and play it back

        If ScriptingChoice$ = "P" Then
            ScriptContents$ = _ReadFile$("WIM_SCRIPT.TXT")
            Kill ("WIM_SCRIPT.TXT")
            _ScreenPrint ScriptContents$
        End If
        EiCfgHandling
        GoTo InjectUpdates
    Case 4
        GoTo BypassWin11Requirements
    Case 5
        GoTo MakeBootDisk
    Case 6
        ExcludeAutounattend$ = "Y"
        EiCfgHandling
        GoTo MakeMultiBootImage
    Case 7
        GoTo MakeBootDisk2
    Case 8
        GoTo ChangeOrder
    Case 9
        GoTo ConvertEsdOrWim
    Case 10
        GoTo GetWimInfo
    Case 11
        GoTo NameAndDescription
    Case 12
        GoTo ExportDrivers
    Case 13
        GoTo ExpandDrivers
    Case 14
        GoTo CreateVHDX
    Case 15
        GoTo AddVHDtoBootMenu
    Case 16
        GoTo CreateISOImage
    Case 17
        GoTo GetFolderToClean
    Case 18
        GoTo AnswerFileGen
    Case 19
        GoTo ProgramHelp
    Case 20
        GoTo ProgramEnd
End Select

' We arrive here if the user makes an invalid selection from the main menu

Cls
Color 14, 4
Print
Print "You have made an invalid selection.";
Color 15
Print " You need to make a selection by entering a number from 1 to 20."
Pause
GoTo MainMenu


' ******************************************************************************************************
' * Inject Windows updates into one or more Windows editions and create a multi edition bootable image *
' ******************************************************************************************************
' AND
' **********************************************************************************************
' * Inject drivers into one or more Windows editions and create a multi edition bootable image *
' **********************************************************************************************
' AND
' ************************************************************************************************************
' * Inject boot-critical drivers into one or more Windows editions and create a multi edition bootable image *
' ************************************************************************************************************


InjectUpdates:

' This routine will inject Windows updates or drivers into one or more Windows ISO images. The ISO images can be
' clean, plain Windows ISO images, or they can be ISO images with an answer file for unattended setup,
' or even ISO images with a sysprep installation.

' Ask for source folder. Check to make sure folder contains ISO images. If it does, ask if all ISO images should be processed.
' For each image to be processed, we need to keep track of the image name to be processed. Likewise, we need to track source folder.

' The variable InjectionMode$ will be set to "UPDATES" if we are injecting Windows updates, and "DRIVERS" if we are injecting drivers

' Initialize variables

DISM_Error_Found$ = ""
SourceFolder$ = ""
FileCount = 0
TotalFiles = 0
ReDim UpdateFlag(0) As String
ReDim FileArray(0) As String
ReDim FolderArray(0) As String ' A list of folders containing files to be processed. Note that the folder path stored here will end with a "\"
DISM_Error_Found$ = "N"
OpsPending$ = "N"

GetFolders:

Do
    Cls
    Print "Enter the path to one or more Windows ISO image files. These can be x64, x86, or dual architecture. These images must"
    Print "include install.wim file(s) ";: Color 0, 10: Print "NOT";: Color 15: Print " install.esd. "
    Print
    Print "NOTE: You can specify a path and we will prompt you regarding each ISO image found there or you can specify a full path"
    Print "with an ISO image file name. ISO image file names ";: Color 0, 10: Print "MUST";: Color 15: Print " end with a .ISO file extension."
    Print
    Line Input "Enter the path: ", SourceFolder$

    If ScriptingChoice$ = "R" Then
        Print #5, ":: Path to one or more Windows images or full path with a file name:"
        If SourceFolder$ = "" Then
            Print #5, "<ENTER>"
        Else
            Print #5, SourceFolder$
        End If
        Print #5, ""
    End If

    CleanPath SourceFolder$
    SourceFolder$ = Temp$

Loop While SourceFolder$ = ""

' Determine if the source is a file name or a folder name.


If _DirExists(SourceFolder$) Then
    ' The name specified is a legit folder name
    SourceFolderIsAFile$ = "N"
    GoTo FolderNameOK
End If

If UCase$(Right$(SourceFolder$, 4)) = ".ISO" Then

    If _FileExists(SourceFolder$) Then
        SourceFolderIsAFile$ = "Y"
        GoTo FolderNameOK
    Else

        ' The name specified was not a valid file or path name

        Cls
        Print "You did not specify a valid folder or file name. Please check the name and try again."

        If ScriptingChoice$ = "R" Then
            Print #5, ":: The above response was not a valid folder or file name."
            Print #5, ""
        End If
        Pause

        GoTo GetFolders
    End If
End If

FolderNameOK:

' We arrice here if the name specified was either a valid file or folder name. If SourceFolderIsAFile$="Y" then
' we need to treat this an individual file and not an entire folder that needs to be parsed for valid files.

CleanPath SourceFolder$

' If the source is a path without a file name then we want a trailing "\".

If SourceFolderIsAFile$ = "N" Then
    SourceFolder$ = Temp$ + "\"
End If

' Perform a check to see if files with a .ISO extension exist in specified folder.
' We are going to call the FileTypeSearch subroutine for this. We will pass to it
' the path to search and the extension to search for. It will return to us the number
' of files with that extension in the variable called filecount and the name of
' each file in an array called FileArray$(x).

' If the source specified was a single file rather than a folder name, we need to make exceptions for this condition.

If SourceFolderIsAFile$ = "N" Then
    FileTypeSearch SourceFolder$, ".ISO", "N"

    ' FileTypeSearch returns number of ISO images found as NumberOfFiles and each of those files as TempArray$(x)

    FileCount = NumberOfFiles
Else

    ' Since this is a single file name, we pretend that we searched for files and found just one.

    FileCount = 1
    TempArray$(1) = SourceFolder$

End If

Cls

If FileCount = 0 Then
    Print
    Color 14, 4
    Print "No files with the .ISO extension were found.";
    Color 15
    Print " Please specify another folder."
    If ScriptingChoice$ = "R" Then
        Print #5, ":: No .ISO files were located at the specified location. Another location will need to be selected."
        Print #5, ""
    End If
    Pause
    Cls
    GoTo GetFolders
End If

' If we arrive here, then files with a .ISO extension were found at the location specified.
' FileCount holds the number of .ISO files found.

UpdateAll$ = "N" ' Set an initial value

UpdateAll:

' If there is only 1 file, then automatically set the UpdateAll$ flag to "Y", otherwise, ask if all files in folder should be updated.

Cls

If FileCount > 1 Then
    Print "Do you want to update at least one Windows edition from ";: Color 0, 10: Print "ALL";: Color 15: Print " of the files located here";: Input UpdateAll$
    If ScriptingChoice$ = "R" Then
        Print #5, ":: Do you want to update at least one Windows edition from ALL of the files located in the specified folder?"
        If UpdateAll$ = "" Then
            Print #5, "<ENTER>"
        Else
            Print #5, UpdateAll$
        End If
        Print #5, ""
    End If
Else
    UpdateAll$ = "Y"
End If

YesOrNo UpdateAll$
UpdateAll$ = YN$
If UpdateAll$ = "X" Then GoTo UpdateAll

If UpdateAll$ = "Y" Then

    For x = 1 To FileCount
        TotalFiles = TotalFiles + 1

        ' Init variables

        ReDim _Preserve UpdateFlag(TotalFiles) As String
        ReDim _Preserve FileArray(TotalFiles) As String
        ReDim _Preserve FolderArray(TotalFiles) As String
        ReDim _Preserve FileSourceType(TotalFiles) As String

        UpdateFlag$(TotalFiles) = "Y"
        FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
        FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
        Cls
        Print "Please standby for a moment. Verifying the following image:"
        Print
        Color 10
        Print FileArray$(TotalFiles)
        Color 15
        Print
        Temp$ = FolderArray$(TotalFiles) + FileArray$(TotalFiles)
        DetermineArchitecture Temp$, 1
        Select Case ImageArchitecture$
            Case "x64"
                FileSourceType$(TotalFiles) = ImageArchitecture$
            Case "DUAL", "NONE", "x86"
                Cls
                Print
                Color 14, 4: Print "WARNING!";: Color 15: Print " An invalid file has been selected."
                Print
                Print "Check the following file to make sure that it is valid. It needs to contain an install.wim file, not INSTALL.ESD."
                Print "In addition, make sure that all files are valid x64 Windows image files."
                Print
                Print "Path: ";: Color 10: Print Left$(Temp$, ((_InStrRev(Temp$, "\"))) - 1): Color 15
                Print "File: ";: Color 10: Print Right$(Temp$, (Len(Temp$) - (_InStrRev(Temp$, "\")))): Color 15
                If ScriptingChoice$ = "R" Then
                    Print #5, ":: An invalid file was selected."
                    Print #5, ""
                End If
                Pause
                ChDir ProgramStartDir$: GoTo BeginProgram
        End Select
    Next x
    GoTo CheckForMoreFolders
End If

' We end up here if the user does NOT want to update every image in the selected location.
' In that case, we need to ask the user about each image file to see if it contains any
' Windows editions that the user wants to have updated.

For x = 1 To FileCount

    Marker1:

    Cls
    Print "Do you want to update any of the Windows editions in this file?"
    Print
    Color 4: Print "Location:  ";: Color 10: Print Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
    Color 4: Print "File name: ";: Color 10: Print Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
    Color 15
    Print
    Input "Do you want to update any of the Windows editions in this file"; UpdateThisFile$
    If ScriptingChoice$ = "R" Then
        Print #5, ":: Do you want to update any of the Windows editions in this file?"
        Print #5, "::    Path:  "; Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
        Print #5, "::    File name: "; Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
        If UpdateThisFile$ = "" Then
            Print #5, "<ENTER>"
        Else
            Print #5, UpdateThisFile$
        End If
        Print #5, ""
    End If
    YesOrNo UpdateThisFile$
    Select Case YN$
        Case "X"
            Print
            Color 14, 4
            Print "Please provide a valid response."
            Color 15
            If ScriptingChoice$ = "R" Then
                Print #5, ":: The above response was invalid."
                Print #5, ""
            End If
            Pause
            GoTo Marker1
        Case "Y"
            TotalFiles = TotalFiles + 1

            ' Check validity of any selected files. The file should be a valid x64 Windows image file.

            ' Init variables

            ReDim _Preserve UpdateFlag(TotalFiles) As String
            ReDim _Preserve FileArray(TotalFiles) As String
            ReDim _Preserve FolderArray(TotalFiles) As String
            ReDim _Preserve FileSourceType(TotalFiles) As String

            UpdateFlag$(TotalFiles) = "Y"
            FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
            FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
            Cls
            Print "Please standby for a moment. Verifying the the following image:"
            Print
            Color 10
            Print FileArray$(TotalFiles)
            Color 15
            Print
            Temp$ = FolderArray$(TotalFiles) + FileArray$(TotalFiles)
            DetermineArchitecture Temp$, 1
            Select Case ImageArchitecture$
                Case "x64"
                    FileSourceType$(TotalFiles) = ImageArchitecture$
                Case "Dual", "NONE", "x86"
                    Cls
                    Print
                    Color 14, 4: Print "WARNING!";: Color 15: Print " An invalid file has been selected."
                    Print
                    Print "Check the following file to make sure that it is valid. It needs to contain an install.wim file, not INSTALL.ESD."
                    Print "In addition, make sure that all files are valid x64 Windows image files."
                    Print
                    Print "Path: ";: Color 10: Print Left$(Temp$, ((_InStrRev(Temp$, "\"))) - 1): Color 15
                    Print "File: ";: Color 10: Print Right$(Temp$, (Len(Temp$) - (_InStrRev(Temp$, "\")))): Color 15
                    If ScriptingChoice$ = "R" Then
                        Print #5, ":: The following file is not valid:"
                        Print #5, "::    Path: "; Left$(Temp$, ((_InStrRev(Temp$, "\"))) - 1)
                        Print #5, "::    File: "; Right$(Temp$, (Len(Temp$) - (_InStrRev(Temp$, "\"))))
                        Print #5, ""
                    End If
                    Pause
                    ChDir ProgramStartDir$: GoTo BeginProgram
            End Select
    End Select
Next x

GoTo CheckForMoreFolders

CheckForMoreFolders:

MoreFolders$ = "" ' Initial value
Cls
Input "Do you want to specify another file name or folder with more ISO images"; MoreFolders$

If ScriptingChoice$ = "R" Then
    Print #5, ":: Do you want to specify another file name or folder with more ISO images?"
    If MoreFolders$ = "" Then
        Print #5, "<ENTER>"
    Else
        Print #5, MoreFolders$
    End If
    Print #5, ""
End If

YesOrNo MoreFolders$

Select Case YN$
    Case "X"
        Print
        Color 14, 4
        Print "Please provide a valid response."
        Color 15
        If ScriptingChoice$ = "R" Then
            Print #5, ":: The previous response was invalid."
            Print #5, ""
        End If
        Pause
        GoTo CheckForMoreFolders
    Case "Y"
        Cls
        GoTo GetFolders
End Select

' At this point, we have a list of all the folders and ISO image files that have Windows editions that are
' to be updated. First, we are going to verify that at least one file has been selected for update. If not,
' go back to the main menu.

If TotalFiles = 0 Then
    Cls
    Color 14, 4: Print "You have not selected any files to use.";: Color 15: Print " We will now return to the main menu."

    ' The user did not select any files to use. We are aborting and returning to the start of program.
    ' Since the script being recorded is not valid, we will close it and delete it.

    If ScriptingChoice$ = "R" Then
        Close #5
        Kill "WIM_SCRIPT.TXT"
    End If
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

GetIndexList:

ReDim IndexCount(TotalFiles) As Integer ' Initialize array

For IndexCountLoop = 1 To TotalFiles
    If UpdateFlag$(IndexCountLoop) = "N" Then GoTo NoIndex

    GetMyIndexList:

    Cls
    Print "Enter the index number(s) for the Windows editions to be updated: "
    Print
    Color 4: Print "File Location: ";: Color 10: Print FolderArray$(IndexCountLoop)
    Color 4: Print "File Name    : ";: Color 10: Print FileArray$(IndexCountLoop)
    Color 15
    Print
    Print "> To view a list of available indices: Press ENTER."
    Print "> To view help for entering index numbers: Type HELP and press ENTER."
    Print "Otherwise, enter the index number(s) and press ENTER."
    Print
    Input "Enter index number(s), press ENTER, or type HELP and press ENTER: ", IndexRange$
    IndexRange$ = UCase$(IndexRange$)

    If ScriptingChoice$ = "R" Then
        Print #5, ":: For this file..."
        Print #5, "::    "; FolderArray$(IndexCountLoop)
        Print #5, "::    "; FileArray$(IndexCountLoop)
    End If


    If Left$(IndexRange$, 1) = "H" Then
        Cls
        Print "You can enter a single index number or multiple index numbers. To enter a contiguous range of index numbers,"
        Print "separate the numbers with a dash like this: 1-4. For non contiguous indices, separate them with a space like"
        Print "this: 1 3. You can also combine both methods like this: 1-3 5 7-9. Make sure to enter numbers from low to high."
        Print
        Print "Finally, if you want to add all editions of Windows, simply enter "; Chr$(34); "ALL"; Chr$(34); "."

        If ScriptingChoice$ = "R" Then
            Print #5, ":: Rather than specifying an index number, you asked for HELP."
            Print #5, "HELP"
            Print #5, ""
        End If

        Pause

        GoTo GetMyIndexList

    End If

    ' We arrive here if help was NOT requested

    If ScriptingChoice$ = "R" Then
        If IndexRange$ = "" Then
            Print #5, ":: A listing of available indices was requested by pressing <ENTER>."
            Print #5, "<ENTER>"
            Print #5, ""
        Else
            Print #5, ":: These indices were specified:"
            Print #5, IndexRange$
            Print #5, ""
        End If
    End If

    If ((IndexRange$ <> "") And (IndexRange$ <> "ALL")) Then GoTo ProcessRange
    SourcePath$ = FolderArray$(IndexCountLoop) + FileArray$(IndexCountLoop)

    Select Case IndexRange$
        Case ""
            Silent$ = "N"
        Case "ALL"
            Silent$ = "Y"
    End Select

    GoSub DisplayIndices2
    If IndexRange$ = "" Then GoTo GetMyIndexList
    If IndexRange$ = "ALL" Then
        Temp$ = ""
        GetNumberOfIndices
        Temp$ = _Trim$(Str$(NumberOfSingleIndices))
        IndexRange$ = "1-" + Temp$
        If IndexRange$ = "1-1" Then IndexRange$ = "1"
    End If
    Kill "Image_Info.txt"

    ProcessRange:

    ProcessRangeOfNums IndexRange$, 1
    If ValidRange = 0 Then
        Color 14, 4
        Print "You did not enter a valid range of numbers"
        Color 15
        If ScriptingChoice$ = "R" Then
            Print #5, ":: The range of indices specified was not valid."
            Print #5, ""
        End If
        Pause
        GoTo GetMyIndexList
    End If

    ' We will now get image info and save it to a file called Image_Info.txt. We will parse that file to verify that the index
    ' selected is valid. If not, we will ask the user to choose a valid index.

    SourcePath$ = FolderArray$(IndexCountLoop) + FileArray$(IndexCountLoop)
    Print
    Print "Verifying indices."
    Print
    Print "Please standby..."
    Print
    GetWimInfo_Main SourcePath$, 1

    For x = 1 To TotalNumsInArray
        WimInfoFound = 0 ' Init Variable
        Open "Image_Info.txt" For Input As #1
        Do
            Line Input #1, WimInfo$
            If Len(WimInfo$) >= 9 Then
                If (Left$(WimInfo$, 7) = "Index :") And (Val(Right$(WimInfo$, (Len(WimInfo$) - 8))) = RangeArray(x)) Then
                    Line Input #1, WimInfo$
                    NameFromFile$ = Right$(WimInfo$, (Len(WimInfo$) - 7))
                    Line Input #1, WimInfo$
                    DescriptionFromFile$ = Right$(WimInfo$, (Len(WimInfo$) - 14))
                    WimInfoFound = 1
                End If
            End If

            SkipToNextLine_Section1:

        Loop Until EOF(1)
        Close #1
        If WimInfoFound = 0 Then
            Cls
            Color 14, 4
            Print "Index"; RangeArray(x); "was not found."
            Print "Please supply a valid index number."
            Color 15
            If ScriptingChoice$ = "R" Then
                Print #5, ":: An invalid index number was specified."
                Print #5, ""
            End If
            Pause
            GoTo GetMyIndexList
        End If
        IndexCount(IndexCountLoop) = TotalNumsInArray

        ' For the index list, we are making an assumption that there will never be more than 100 indicies in an image.

        ReDim _Preserve IndexList(TotalFiles, 100) As Integer

        For y = 1 To TotalNumsInArray
            IndexList(IndexCountLoop, y) = RangeArray(y)
        Next y
    Next x
    Kill "Image_Info.txt"

    NoIndex:

Next IndexCountLoop

' Now that we have a valid source directory and we know that there are ISO images
' located there, ask the user for the location where we should save the updated files.

DestinationFolder$ = "" ' Set initial value

GetDestinationPath10:

Do
    Cls
    Print "Enter the path where the project will be created. This is where all the temporary files will be stored and we will"
    Print "save the final ISO image file here as well."
    Print
    Line Input "Enter the path where the project should be created: ", DestinationFolder$

    If ScriptingChoice$ = "R" Then
        Print #5, ":: Enter the path where the project will be created:"
        If DestinationFolder$ = "" Then
            Print #5, "<ENTER>"
        Else
            Print #5, DestinationFolder$
        End If
        Print #5, ""
    End If
Loop While DestinationFolder$ = ""

CleanPath DestinationFolder$
DestinationFolder$ = Temp$ + "\"

' We don't want user to specify the root of a drive

If Len(DestinationFolder$) = 3 Then
    Cls
    Color 14, 4
    Print "Please do not specify the root directory of a drive."
    Color 15
    If ScriptingChoice$ = "R" Then
        Print #5, ":: It appears that the root directory of a drive was specified. This is not a valid location."
        Print #5, ""
    End If
    Pause
    GoTo GetDestinationPath10
End If

' Check to see if the destination specified is on a removable disk

Cls
Print "Performing a check to see if the destination you specified is a removable disk."
Print
Print "Please standby..."
DriveLetter$ = Left$(DestinationFolder$, 2)
RemovableDiskCheck DriveLetter$
DestinationIsRemovable = IsRemovable

Select Case DestinationIsRemovable
    Case 2
        Cls
        Color 14, 4: Print "This is not a valid disk.";: Color 15: Print " Please specify another location."
        If ScriptingChoice$ = "R" Then
            Print #5, ":: An invalid disk was specified."
            Print #5, ""
        End If
        Pause
        GoTo GetDestinationPath10
    Case 1
        Cls
        Print "The disk that you specified is a removable disk. ";: Color 14, 4: Print "Please specify a fixed disk.": Color 15
        Print
        Print "NOTE: Project must be created on a fixed disk due to limitations of some Microsoft utilities."
        If ScriptingChoice$ = "R" Then
            Print #5, ":: The specified disk is a removable disk. This is not valid."
            Print #5, ""
        End If
        Pause
        GoTo GetDestinationPath10
    Case 0
        ' if the returned value was a 0, no action is necessary. The program will continue normally.
End Select

' Verify that the path specified exists.

If Not (_DirExists(DestinationFolder$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(DestinationFolder$)) Then
        Cls
        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
        Print
        Print "Please recheck the path you have specified and try again."
        If ScriptingChoice$ = "R" Then
            Print #5, ":: The destination does not exist and could not be created."
            Print #5, ""
        End If
        Pause
        GoTo GetDestinationPath10
    End If
End If

' If we have arrived here it means that the destination path already exists
' or we were able to create it successfully.

' Start by setting an AV exclusion for the destination path. We will log this location to a temporary file
' so that if the file is interrupted unexpectedly, we can remove the exclusion the next time the program
' is started.

' IMPORTANT: The count of files listed immediately below is the number of files of each type in the folders specified
' INCLUDING FILES THAT WILL NOT BE UPDATED.

' Add an AV exclusion for the destination folder

CleanPath DestinationFolder$
ff = FreeFile
Open "WIM_Exclude_Path.txt" For Output As #ff
Print #ff, Temp$
Close #ff
Cmd$ = "powershell.exe -command Add-MpPreference -ExclusionPath " + "'" + Chr$(34) + Temp$ + Chr$(34) + "'"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

' The next set of variables will hold the actual number of each image type to be processed

TotalImagesToUpdate = 0

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then TotalImagesToUpdate = TotalImagesToUpdate + IndexCount(x)
Next x

' Get the location for updates or drivers to be injected

GetUpdatesLocation:

x64Updates$ = "" 'Set initial value

If InjectionMode$ = "UPDATES" Then
    Do
        Cls
        Print "Please enter the path to the Windows update files. Please specify only the base path to these files without the x64 or"
        Print "x86 folders. The program will automatically determine whether it needs to select files in the x64 or x86 folder."
        Print
        Line Input "Enter the path to the Windows update files: ", UpdatesLocation$
        If ScriptingChoice$ = "R" Then
            Print #5, ":: Enter the path to the Windows update files:"
            If UpdatesLocation$ = "" Then
                Print #5, "<ENTER>"
            Else
                Print #5, UpdatesLocation$
            End If
            Print #5, ""
        End If
    Loop While UpdatesLocation$ = ""
End If

If InjectionMode$ = "DRIVERS" Then
    Do
        Cls
        Line Input "Enter the path to the drivers: ", UpdatesLocation$
        If ScriptingChoice$ = "R" Then
            Print #5, ":: Enter the path to the drivers:"
            If UpdatesLocation$ = "" Then
                Print #5, "<ENTER>"
            Else
                Print #5, UpdatesLocation$
            End If
            Print #5, ""
        End If
    Loop While UpdatesLocation$ = ""
End If

If InjectionMode$ = "BCD" Then
    Do
        Cls
        Line Input "Enter the path to the boot-critical drivers: ", UpdatesLocation$
        If ScriptingChoice$ = "R" Then
            Print #5, ":: Enter the path to the boot-critical drivers:"
            If UpdatesLocation$ = "" Then
                Print #5, "<ENTER>"
            Else
                Print #5, UpdatesLocation$
            End If
            Print #5, ""
        End If
    Loop While UpdatesLocation$ = ""
End If

CleanPath UpdatesLocation$
UpdatesLocation$ = Temp$

x64Updates$ = UpdatesLocation + "\x64"

If TotalImagesToUpdate = 0 Then GoTo End_Getx64UpdatesLocation

' Verify that the x64 path specified exists.

If Not (_DirExists(x64Updates$)) Then

    ' The path does not exist. Inform user and allow them to try again.

    Cls
    Color 14, 4: Print "The specified x64 folder does not exist.";: Color 15: Print " Please try again."
    Print
    Print "All updates should be located within an x64 subfolder. For example, if you specify that updates are located in"
    Print "D:\WinUpdates, then you should have a folder called D:\WinUpdates\x64 and under that would be all of the subfolders"
    Print "such as LCU, SafeOS_DU, etc."
    If ScriptingChoice$ = "R" Then
        Print #5, ":: The specified x64 folder does not exist."
        Print #5, ""
    End If
    Pause
    GoTo GetUpdatesLocation
End If

' If we have arrived here it means that the path is valid.
' Now, verify that update files or drivers actually exist in this location.

Select Case InjectionMode$
    Case "UPDATES"
        NumberOfx64Updates = 0 ' Set initial value

        FileTypeSearch (x64Updates$ + "\SSU\"), ".MSU", "N"
        NumberOfx64Updates = NumberOfx64Updates + NumberOfFiles

        FileTypeSearch (x64Updates$ + "\LCU\"), ".MSU", "N"
        NumberOfx64Updates = NumberOfx64Updates + NumberOfFiles

        FileTypeSearch (x64Updates$ + "\Other\"), "*", "N"
        NumberOfx64Updates = NumberOfx64Updates + NumberOfFiles

        FileTypeSearch (x64Updates$ + "\Setup_DU\"), ".CAB", "N"
        NumberOfx64Updates = NumberOfx64Updates + NumberOfFiles

        FileTypeSearch (x64Updates$ + "\SafeOS_DU\"), ".CAB", "N"
        NumberOfx64Updates = NumberOfx64Updates + NumberOfFiles

        FileTypeSearch (x64Updates$ + "\PE_Files\"), "*", "N"
        NumberOfx64Updates = NumberOfx64Updates + NumberOfFiles

        If _FileExists(UpdatesLocation$ + "\Answer_File\autounattend.xml") Then
            NumberOfx64Updates = NumberOfx64Updates + 1
            AddAnswerFile$ = "Y"
        Else
            AddAnswerFile$ = "N"
        End If

        If NumberOfx64Updates = 0 Then
            Cls
            Print
            Color 14, 4: Print "No x64 update files were found in this location.": Color 15
            Print
            Print "Please specify another location."
            If ScriptingChoice$ = "R" Then
                Print #5, ":: No x64 update files were found at the location specified."
                Print #5, ""
            End If
            Pause
            GoTo GetUpdatesLocation
        End If
    Case "DRIVERS", "BCD"
        FileTypeSearch x64Updates$ + "\", ".INF", "Y"
        NumberOfx64Updates = NumberOfFiles

        If NumberOfx64Updates = 0 Then
            Cls
            Print
            Color 14, 4: Print "No x64 drivers were found in this location.": Color 15
            Print
            Print "Please specify another location."
            If ScriptingChoice$ = "R" Then
                Print #5, ":: No x64 drivers were found at the location specified."
                Print #5, ""
            End If
            Pause
            GoTo GetUpdatesLocation
        End If
End Select

End_Getx64UpdatesLocation:

' Ask user what they want to name the final ISO image file

Cls
UserSelectedImageName$ = "" ' Set initial value
Print "If you would like to specify a name for the final ISO image file that this project will create, please do so now,"
Print "WITHOUT an extension. You can also simply press ENTER to use the default name of Windows.ISO."
Print
Print "Enter name ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension, or press ENTER: ";: Line Input "", UserSelectedImageName$

If UserSelectedImageName = "" Then
    UserSelectedImageName$ = "Windows.iso"
Else
    UserSelectedImageName$ = UserSelectedImageName$ + ".iso"
End If

If ScriptingChoice$ = "R" Then
    Print #5, ":: Name of final ISO image (without an extension):"
    Print #5, Left$(UserSelectedImageName$, _InStrRev(UserSelectedImageName$, ".") - 1)
    Print #5, ""
End If

' If the user was recording a script, we now have all the information needed. The script will now
' be saved and we will ask the user if they want to continue the update process or just save the
' recorded script and stop.

If ScriptingChoice$ = "R" Then
    Print #5, "::::::::::::::::::::::::::::::::::::::::::::::"
    Print #5, ":: Script created on "; Date$; " at "; Time$; " ::"
    Print #5, "::             END OF SCRIPT                ::"
    Print #5, "::::::::::::::::::::::::::::::::::::::::::::::"
    Close #5

    ' Turn off script recording

    ScriptingChoice$ = "S"

    ' Move the newly recorded script to the program folder.

    Cmd$ = "move /y WIM_SCRIPT.TXT " + Chr$(34) + ProgramStartDir$ + Chr$(34)
    Shell _Hide Cmd$

    ContinueWithUpdates:

    Cls
    Print "The creation of the script file has now been completed. If your intention was to simply create a script file we can end"
    Print "the program now. If you would like to continue with the updates, we can now perform the update operations that you"
    Print "have requested."
    Print
    Input "Do you want to continue and perform the updates now? ", Temp$
    YesOrNo Temp$
    Temp$ = YN$

    If Temp$ = "X" Then GoTo ContinueWithUpdates

    If Temp$ = "N" Then ChDir ProgramStartDir$: GoTo BeginProgram

End If

' All the information needed from the user has been gathered. If a script was being recorded, it has been saved.
' Also, if a script was being recorded, we arrive here if the user wanted to continue on and apply the updates
' that they requested and not simply stop with the creation of the script.

' If there is an old logs folder made before this project, delete it, then create a new, empty logs folder
' so we start clean with this project.

' The logs folder is different than the other folders we will use in the project because we want to keep
' it through each loop for each file that we process. As a result, we create it before the main loop
' through each file to process is started.

Cls
Print "Please standby while we do a little housekeeping. Note that this can possibly take a while."

' Cleaning up previous stale DISM operations

x = 0
Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /get-mountedimageinfo > DismInfo.txt"
Shell Chr$(34) + Cmd$ + Chr$(34)
Open "DismInfo.txt" For Input As #1

Do
    x = x + 1
    Line Input #1, (MountDir$)
    If (InStr(MountDir$, "Mount Dir :")) Then
        If x = 1 Then
            Print
            Color 14, 4: Print "Warning!";: Color 15: Print " There is still at least one image mounted by DISM open."
            Print "We will try to clear these at this time. Please standby..."
        End If
        MountDir$ = Right$(MountDir$, (Len(MountDir$) - 12))
        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + MountDir$ + Chr$(34) + " /Discard"
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Cleanup-WIM"
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /cleanup-mountpoints"
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    End If
Loop Until EOF(1)

Close #1
Kill "disminfo.txt"

' Run a check a second time for open mounts. The above procedure should have cleared any open mounts, but if we still
' have an open mount then we may need to reboot to resolve the situation.

Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /get-mountedimageinfo > DismInfo.txt"
Shell Chr$(34) + Cmd$ + Chr$(34)
Open "DismInfo.txt" For Input As #1

Do
    Line Input #1, MountDir$
    If (InStr(MountDir$, "Mount Dir :")) Then
        Close #1
        Cls
        Print "We were not able to clear all mounts. Try rebooting the system and then run the program again."
        Pause
        Kill "disminfo.txt"
        ChDir ProgramStartDir$: GoTo BeginProgram
    End If
Loop Until EOF(1)

Close #1
Kill "disminfo.txt"

' Cleanup logs folder

TempPath$ = DestinationFolder$ + "logs\"

If _DirExists(TempPath$) Then
    Cmd$ = "rmdir " + Chr$(34) + TempPath$ + Chr$(34) + " /s /q"
    Shell _Hide Cmd$
End If

If Not (_DirExists(TempPath$)) Then
    MkDir DestinationFolder$ + "logs"
End If

' Before starting the update process, verify that there are no leftover files sitting in the
' destination.

Cleanup DestinationFolder$
If CleanupSuccess = 0 Then ChDir ProgramStartDir$: GoTo BeginProgram

' Create the folders we need for the project.

MkDir DestinationFolder$ + "Mount"
MkDir DestinationFolder$ + "ISO_Files"
MkDir DestinationFolder$ + "Scratch"
MkDir DestinationFolder$ + "Temp"
MkDir DestinationFolder$ + "WIM_x64"
MkDir DestinationFolder$ + "Assets"
MkDir DestinationFolder$ + "WinRE_MOUNT"
MkDir DestinationFolder$ + "WinPE_MOUNT"
MkDir DestinationFolder$ + "WinPE"
MkDir DestinationFolder$ + "WinRE"
MkDir DestinationFolder$ + "Setup_DU_x64"
MkDir DestinationFolder$ + "SSU_x64"


' Prior to starting the exports, we need to initialize some variables. To init these variables, we need to know how many indices we will be handling
' so we will determine that now.

TotalIndexCount = 0 'Init variable

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then
        TotalIndexCount = TotalIndexCount + IndexCount(x)
    End If
Next x

ReDim x64OriginalFile(0) As String
ReDim x64SourceArc(0) As String
ReDim x64OriginalIndex(0) As String
CurrentIndexCount = 0
x64ExportCount = 0

' We are going to create a new PendingOps.log file. Delete old file if it exists.

If _FileExists(DestinationFolder$ + "logs\PendingOps.log") Then Kill DestinationFolder$ + "logs\PendingOps.log"

ff = FreeFile
Open (DestinationFolder$ + "logs\PendingOps.log") For Output As #ff
Print #ff, "Below is a list of files that were found with pending operations. Pending operations are the result of items added to the image that will"
Print #ff, "prevent DISM from being able to perform a cleanup operation on the image. The most common of the causes is enabling NetFX3. If this log"
Print #ff, "file lists any files below that have pending operations, you should redo this update using source files that do not have any pending"
Print #ff, "operations present."
Print #ff, "-------------------------------------------------------------------------------------------------------------------------------------------"
Print #ff, ""
Close #ff

Select Case InjectionMode$
    Case "UPDATES"
        AddUpdatesStatusDisplay 0, 0, 1
    Case "DRIVERS"
        AddUpdatesStatusDisplay 0, 0, 16
    Case "BCD"
        AddUpdatesStatusDisplay 0, 0, 50
End Select

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        MountISO Temp$
        For y = 1 To IndexCount(x)
            CurrentIndexCount = CurrentIndexCount + 1
            x64ExportCount = x64ExportCount + 1
            ReDim _Preserve x64OriginalFile(x64ExportCount) As String
            ReDim _Preserve x64SourceArc(x64ExportCount) As String
            ReDim _Preserve x64OriginalIndex(x64ExportCount) As String
            x64OriginalFile$(x64ExportCount) = Temp$
            x64SourceArc$(x64ExportCount) = "x64"
            x64OriginalIndex$(x64ExportCount) = LTrim$(Str$(IndexList(x, y)))
            SourceArcFlag$ = ""
            DestArcFlag$ = "WIM_x64"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + SourceArcFlag$ + "\sources" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\WinPE"_
                    + CHR$(34) + " boot.wim /A-:RHS > NUL"
            Shell _Hide Cmd$
            Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\WinPE\boot.wim" + Chr$(34) + " BOOT_x64.wim"
            Shell _Hide Cmd$
            CurrentIndex$ = LTrim$(Str$(IndexList(x, y)))
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + MountedImageDriveLetter$ + SourceArcFlag$_
            + "\Sources\install.wim" + CHR$(34) + " /SourceIndex:" + CurrentIndex$ + " /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + DestArcFlag$_
            + "\install.wim" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Next y

        ' The next command dismounts the ISO image since we are now done with it. The messages displayed by the process are
        ' not really helpful so we are going to hide those messages even if detailed status is selected by the user.

        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    End If
Next x

' At this point, all images have been exported from their original files to the project folder \WIM_x64.
' Now we need to mount the install.wim file and update all the images therein.

' Ditch the trailing backslash (robocopy does not like it)

CleanPath DestinationFolder$
DestinationFolder$ = Temp$

' The following section is run for editions that are having Windows updates injected

' We begin by updating the WinRE.wim and the Boot.wim (for WinPE). We only need to process the WinRE and boot.wim once
' so we will do that on the first edition that we process.

CurrentImage = 0

If (TotalImagesToUpdate > 0 And InjectionMode$ = "UPDATES") Then
    For x = 1 To TotalImagesToUpdate
        CurrentImage = CurrentImage + 1
        Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
        CurrentIndex$ = LTrim$(Str$(x))
        If x = 1 Then
            AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 2
        Else
            AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 25
        End If
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WIM_x64\install.wim" + CHR$(34)_
        + " /Index:" + CurrentIndex$ + " /Mountdir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Check the current Windows edition to see if it has any pending operations.

        Cmd$ = "PowerShell " + Chr$(34) + "Get-WindowsCapability -path '" + DestinationFolder$ + "\mount" + "' | Where-Object { $_.State -eq 'InstallPending' }" + Chr$(34) + " > capabilities.txt"
        Shell Cmd$
        OpsPendingFileCheck$ = _ReadFile$("capabilities.txt")
        If InStr(OpsPendingFileCheck$, "InstallPending") Then
            OpsPending$ = "Y"
            ff2 = FreeFile
            Open (DestinationFolder$ + "\logs\PendingOps.log") For Append As #ff2
            Print #ff2, "File name: "; x64OriginalFile$(x)
            Print #ff2, "Architecture type: "; x64SourceArc$(x)
            Print #ff2, "Index number: "; x64OriginalIndex$(x)
            Print #ff2, ""
            Close #ff2
        End If

        Kill "capabilities.txt"

        ' Skip WinRE and WinPE if these have already been processed

        If x > 1 Then
            GoTo SkipWinPEx64
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 3
        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\mount\Windows\System32\Recovery" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\WINRE"_
        + CHR$(34) + " winre.wim /A-:RSH > NUL"
        Shell _Hide Cmd$

        ' Mount the WinRE Image if updates are available

        ' Check for the presence of a Stanalone SSU update. There can be rare cases where a standalone SSU is released, so this section handles that.

        FileTypeSearch x64Updates$ + "\SSU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            SASSU_Update_Avail$ = "Y"
        Else
            SASSU_Update_Avail$ = "N"
        End If

        ' Note that the SSU update is now combined with the LCU update. Since these are no longer separate updates, we need to search for the LCU update
        ' since it could potentially include an SSU update.

        FileTypeSearch x64Updates$ + "\LCU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            LCU_Update_Avail$ = "Y"
        Else
            LCU_Update_Avail$ = "N"
        End If

        FileTypeSearch x64Updates$ + "\SafeOS_DU\", ".CAB", "N"

        If NumberOfFiles > 0 Then
            SafeOS_DU_Avail$ = "Y"
        Else
            SafeOS_DU_Avail$ = "N"
        End If

        If (SASSU_Update_Avail$ = "N") And (LCU_Update_Avail$ = "N") And (SafeOS_DU_Avail$ = "N") Then GoTo Skip_WINRE_Update_x64

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\winre.wim" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Apply the Standalone SSU if present

        If SASSU_Update_Avail$ = "Y" Then

            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /PackagePath="_
            + CHR$(34) + x64Updates$ + "\SSU" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        End If

        ' Add SSU Update to WinRE.WIM using the combined LCU / SSU package

        If LCU_Update_Avail$ = "Y" Then

            ' Start by extracting the SSU (if present) from the combined LCU / SSU

            Cmd$ = "expand " + Chr$(34) + x64Updates$ + "\LCU\*.MSU" + Chr$(34) + " /f:SSU*.cab" + " " + Chr$(34) + DestinationFolder$ + "\SSU_x64" + Chr$(34)
            Shell _Hide Cmd$

            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + Destinationfolder$ + "\SSU_x64" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

            ' The following lines that are commented out were previously used when it was necessary to apply the LCU to WinRE to address a security issue.
            ' That procedure is no longer needed.

            '            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /PackagePath="_
            '            + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34)
            '            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        End If

        ' Add SafeOS DU to WinRE.WIM

        If SafeOS_DU_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\SafeOS_DU" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' cleanup image

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Image:" + Chr$(34) + DestinationFolder$ + "\WINRE_MOUNT" + Chr$(34) + " /cleanup-image /StartComponentCleanup" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount image

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + DestinationFolder$ + "\WINRE_MOUNT" + Chr$(34) + " /Commit" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Skip_WINRE_Update_x64:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\winre.wim" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\WINRE_x64.WIM" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' del the temp file

        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\WinRE\winre.wim" + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 1

        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 4

        ' Mount the WinPE Image - Index 1, if Standalone SSU or combined LCU / SSU update exists

        FileTypeSearch x64Updates$ + "\SSU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            SASSU_Update_Avail$ = "Y"
        Else
            SASSU_Update_Avail$ = "N"
        End If

        FileTypeSearch x64Updates$ + "\LCU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            LCU_Update_Avail$ = "Y"
        Else
            LCU_Update_Avail$ = "N"
        End If

        If Skip_PE_Updates$ = "Y" Then GoTo Export_PE_Index1
        If (SASSU_Update_Avail$ = "N") And (LCU_Update_Avail$ = "N") Then GoTo Export_PE_Index1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        If SASSU_Update_Avail$ = "Y" Then

            ' Add the Standalone SSU if present

            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\SSU" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        End If

        ' Add LCU /SSU Update to BOOT.WIM, Index 1

        If LCU_Update_Avail$ = "Y" Then

            ' Add the SSU (if present)

                       Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + Destinationfolder$ + "\SSU_x64" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

            ' Add the LCU

            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        End If

        ' cleanup image

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Image:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /cleanup-image /StartComponentCleanup" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /Commit" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index1:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 2

        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 5

        ' Generic files, such as scripts, that are being added to the boot.wim only need to be added to index 2. For that reason, we search for such
        ' files now, rather than when processing index 1.

        FileTypeSearch x64Updates$ + "\PE_Files\", "*", "N"
        If NumberOfFiles > 0 Then
            PE_Files_Avail$ = "Y"
        Else
            PE_Files_Avail$ = "N"
        End If

        ' Mount the WinPE Image - Index 2, if updates are available

        If (SASSU_Update_Avail$ = "N") And (LCU_Update_Avail$ = "N") And (PE_Files_Avail$ = "N") Then GoTo Export_PE_Index2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /Index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Standalone SSU and combined LCU / SSU Update to BOOT.WIM, Index 2

        If (SASSU_Update_Avail$ = "Y") And (Skip_PE_Updates$ = "N") Then

            ' Add the Standalone SSU

            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\SSU" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        End If

        If (LCU_Update_Avail$ = "Y") And (Skip_PE_Updates$ = "N") Then

            ' Add the SSU (if present)

            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + Destinationfolder$ + "\SSU_x64" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

            ' Add the LCU

            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        End If

        ' Add generic files such as scripts to the boot.wim.

        FileTypeSearch x64Updates$ + "\PE_Files\", "*", "N"

        If NumberOfFiles > 0 Then
            For z = 1 To NumberOfFiles
                If Left$(TempArray$(z), 1) = "-" Then
                    Cmd$ = Right$(TempArray$(z), (Len(TempArray(z)) - _InStrRev(TempArray$(z), "\")))
                    Cmd$ = DestinationFolder$ + "\WINPE_MOUNT\" + Cmd$
                    If _FileExists(Cmd$) Then
                        Kill Cmd$
                    End If
                Else
                    Cmd$ = "Copy " + Chr$(34) + TempArray$(z) + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34)
                    Shell _Hide Cmd$
                End If
            Next z
        End If

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
        + " /cleanup-image /StartComponentCleanup" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /Commit" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index2:

        ' export index 2

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Export-Image /Bootable /SourceImageFile:" + Chr$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + Chr$(34) + " /SourceIndex:2 /DestinationImageFile:" + Chr$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' The WinRE and WinPE components have been updated. We will now proceed with updating of the main OS (install.wim)

        SkipWinPEx64:

        ' Add the Standalone SSU

        If SASSU_Update_Avail$ = "Y" Then

            Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Add-Package /Image:" + Chr$(34) + DestinationFolder$ + "\mount" + Chr$(34) + " /PackagePath=" + Chr$(34) + x64Updates$ + "\SSU" + Chr$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        End If

        ' Add combined LCU / SSU update to main OS (install.wim) if available.

        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 7

        If LCU_Update_Avail$ = "Y" Then

            ' Add the SSU (if present)

            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + Destinationfolder$ + "\SSU_x64" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

            ' Add the LCU

            Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Add-Package /Image:" + Chr$(34) + DestinationFolder$ + "\mount" + Chr$(34) + " /PackagePath=" + Chr$(34) + x64Updates$ + "\LCU" + Chr$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        End If

        ' Copy the WinRE.wim back into the main OS (install.wim)

        Cmd$ = "copy " + CHR$(34) + DestinationFolder$ + "\Assets\winre_x64.wim" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\MOUNT\Windows\System32\Recovery\winre.wim" + CHR$(34)
        Shell _Hide Cmd$
        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 8
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Cleanup-Image /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)_
        + " /StartComponentCleanup /ResetBase /ScratchDir:" + CHR$(34) + DestinationFolder$ + "\Temp" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 9

        ' Check to see if other updates are available

        FileTypeSearch x64Updates$ + "\Other\", "*", "N"
        If NumberOfFiles > 0 Then
            Other_Updates_Avail$ = "Y"
        Else
            Other_Updates_Avail$ = "N"
        End If

        If Other_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\Other" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Looking for and expanding the Setup Dynamic Update if it exists

        FileTypeSearch x64Updates$ + "\Setup_DU\", ".CAB", "N"

        If NumberOfFiles > 0 Then
            Setup_DU$ = TempArray$(1)
            Cmd$ = "expand " + Chr$(34) + Setup_DU$ + Chr$(34) + " -F:* " + Chr$(34) + DestinationFolder$ + "\Setup_DU_x64" + Chr$(34) + " > NUL"
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 10
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Get-Packages /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " > "_
        + CHR$(34) + DestinationFolder$ + "\Logs\x64_" + LTRIM$(STR$(x)) + "_UpdateResults.txt" + CHR$(34) + ""
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 11
        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + DestinationFolder$ + "\mount" + Chr$(34) + " /Commit"
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Next x
End If

' We will now inject boot-critical drivers into the WinRE and WinPE images. We only need to process the WinRE and boot.wim once
' so we will do that on the first edition that we process.

If (TotalImagesToUpdate > 0 And InjectionMode$ = "BCD") Then
    For x = 1 To TotalImagesToUpdate
        CurrentImage = CurrentImage + 1
        Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
        CurrentIndex$ = LTrim$(Str$(x))
        If x = 1 Then
            AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 51
        End If
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WIM_x64\install.wim" + CHR$(34)_
        + " /Index:" + CurrentIndex$ + " /Mountdir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        If x > 1 Then
            GoTo SkipWinPEx64_BCD
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 52
        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\mount\Windows\System32\Recovery" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\WINRE"_
        + CHR$(34) + " winre.wim /A-:RSH > NUL"
        Shell _Hide Cmd$

        ' Mount the WinRE Image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\winre.wim" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Driver /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) +_
         " /Driver:" +chr$(34)+ x64Updates$ + CHR$(34) + " /recurse"  + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' cleanup image

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Image:" + Chr$(34) + DestinationFolder$ + "\WINRE_MOUNT" + Chr$(34) + " /cleanup-image /StartComponentCleanup" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount image

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + DestinationFolder$ + "\WINRE_MOUNT" + Chr$(34) + " /Commit" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Skip_WINRE_Update_x64_BCD:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\winre.wim" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\WINRE_x64.WIM" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' del the temp file

        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\WinRE\winre.wim" + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 1

        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 53

        ' Mount the WinPE Image - Index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Driver /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) +_
         " /Driver:" + chr$(34)+x64Updates$ + CHR$(34) + " /recurse"  + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' cleanup image

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Image:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /cleanup-image /StartComponentCleanup" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /Commit" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index1_BCD:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 2

        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 54

        ' Mount the WinPE Image - Index 2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /Index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Driver /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) +_
         " /Driver:" +chr$(34)+ x64Updates$ + CHR$(34) + " /recurse"  + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
        + " /cleanup-image /StartComponentCleanup"  + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /Commit" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index2_BCD:

        ' export index 2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /Bootable /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /SourceIndex:2 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = "move /y " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x64.wim" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\assets" + Chr$(34)
        Shell _Hide Cmd$
        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x64.wim" + Chr$(34)

        ' The WinRE and WinPE components have been updated. We will now proceed with updating of the main OS (install.wim).

        SkipWinPEx64_BCD:

        Cmd$ = "copy /B " + CHR$(34) + DestinationFolder$ + "\Assets\winre_x64.wim" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\MOUNT\Windows\System32\Recovery\winre.wim" + CHR$(34)
        Shell _Hide Cmd$
        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 55
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Cleanup-Image /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)_
        + " /StartComponentCleanup /ResetBase /ScratchDir:" + CHR$(34) + DestinationFolder$ + "\Temp" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 56
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Get-Packages /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " > "_
        + CHR$(34) + DestinationFolder$ + "\Logs\x64_" + LTRIM$(STR$(x)) + "_UpdateResults.txt" + CHR$(34) + ""
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Unmounting and saving the Windows edition

        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 57
        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + DestinationFolder$ + "\mount" + Chr$(34) + " /Commit" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Next x
End If

' The following section is run for editions that are having drivers injected

If (TotalImagesToUpdate > 0 And InjectionMode$ = "DRIVERS") Then
    For x = 1 To TotalImagesToUpdate
        CurrentImage = CurrentImage + 1
        Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
        CurrentIndex$ = LTrim$(Str$(x))
        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 17
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WIM_x64\install.wim" + CHR$(34)_
        + " /Index:" + CurrentIndex$ + " /Mountdir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 18
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\Mount" + CHR$(34) + " /Add-Driver /Driver:" + CHR$(34)_
        + x64Updates$ + CHR$(34) + " /RECURSE" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 19
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Get-Drivers /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " > " + CHR$(34)_
        + DestinationFolder$ + "\Logs\x64_" + LTRIM$(STR$(x)) + "_UpdateResults.txt" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 20
        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + DestinationFolder$ + "\mount" + Chr$(34) + " /Commit" + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Next x
End If

_ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " by Hannes Sehestedt"

' To ensure that DestinationFolder$ is always specified consistently without a trailing backslash, we will
' run it through the CleanPath routine.

CleanPath DestinationFolder$
DestinationFolder$ = Temp$

' Create a base image

ProjectIsSingleArchitecture:

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        Select Case InjectionMode$
            Case "UPDATES"
                AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 12
            Case "DRIVERS"
                AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 21
            Case "BCD"
                AddUpdatesStatusDisplay CurrentImage, TotalImagesToUpdate, 58
        End Select
        MountISO Temp$
        Select Case ExcludeAutounattend$
            Case "Y"
                            Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
                            + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd autounattend.xml > NUL"
            Case "N"
                            Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
                            + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd > NUL"
        End Select
        Shell _Hide Cmd$
        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Exit For
    End If
Next x

' We need a file called ei.cfg in each "sources" folder. When creating a Windows image that has multiple editions, this file is needed
' to prevent setup from simply installing the version of Windows that originally shipped on a machine without presenting a menu
' to allow you to choose the edition that you want to install. Note that this file is not needed for unattended installs since the
' autounattend.xml answer file will specify which version of Windows to install, but it does not hurt to have the file there.
'
' The following lines will check to see if an ei.cfg file is already present. If so, we will leave it alone in case it is configured
' differenty than what we are going to put in place, otherwise we will create the file.

If CreateEiCfg$ = "Y" Then
    Temp$ = DestinationFolder$ + "\ISO_Files\sources"
    If _DirExists(Temp$) Then
        If Not (_FileExists(Temp$ + "\ei.cfg")) Then
            Open (Temp$ + "\ei.cfg") For Output As #1
            Print #1, "[CHANNEL]"
            Print #1, "Retail"
            Close #1
        End If
    End If
End If

' When we arrive here, the base image for the single architecure type project has been completed.

' Moving the updated install.wim, Boot.wim, and Setup Dynamic Updates to the base image

Select Case InjectionMode$
    Case "UPDATES"
        AddUpdatesStatusDisplay 0, 0, 13
    Case "DRIVERS"
        AddUpdatesStatusDisplay 0, 0, 22
    Case "BCD"
        AddUpdatesStatusDisplay 0, 0, 59
End Select

If TotalImagesToUpdate > 0 Then
    Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
    If _FileExists(Temp$) Then
        Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
        Shell _Hide Cmd$

        For x = 1 To TotalImagesToUpdate
Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Export-Image /SourceImageFile:" + Chr$(34) + DestinationFolder$ + "\WIM_x64\install.wim"_
             + Chr$(34) + " /SourceIndex:" + LTrim$(Str$(x)) + " /DestinationImageFile:" + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources\install.wim"+ Chr$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Next x

            Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM " + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
            + "\ISO_Files\Sources\boot.wim" + CHR$(34) + " > NUL"
        Shell _Hide Cmd$
                        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\Setup_DU_x64" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
                       + "\ISO_Files\Sources" + CHR$(34) + " *.* /e > NUL"
        Shell _Hide Cmd$
    End If
End If

' This code syncs files that are possibly not synced between the WinPE image and the base install.wim image

Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\ISO_Files\Sources\boot.wim" + CHR$(34)_
 + " /index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
Cmd$ = "copy /B /Y " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\Sources\setup.exe" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
Cmd$ = "copy /B /Y " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\Sources\setuphost.exe" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
Cmd$ = "copy /B /Y " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\windows\boot\efi\bootmgfw.efi" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\efi\boot\bootx64.efi" + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
Cmd$ = "copy /B /Y " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\windows\boot\efi\bootmgr.efi" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /Mountdir:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /discard" + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

' End of file syncing.

FinalImageName$ = DestinationFolder$ + "\" + UserSelectedImageName$

' Technical Note: OSCDIMG does not hide its output by simply redirecting to NUL. By using " > NUL 2>&1" we work around this.
' How this works: Standard output is going to NUL and standard error output (file descriptor 2) is being sent to standard output
' (file descriptor 1) so both error and normal output go to the same place.

_ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " - Saving the Final ISO Image"

Select Case InjectionMode$
    Case "UPDATES"
        AddUpdatesStatusDisplay 0, 0, 14

        If AddAnswerFile$ = "Y" Then
            Cmd$ = "robocopy " + Chr$(34) + UpdatesLocation$ + "\Answer_File" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " autounattend.xml /nfl /ndl /njh /njs > nul"
            Shell _Hide Cmd$
        End If

    Case "DRIVERS"
        AddUpdatesStatusDisplay 0, 0, 23
    Case "BCD"
        AddUpdatesStatusDisplay 0, 0, 60
End Select

If _FileExists(DestinationFolder$ + "\ISO_Files\autounattend.xml") Then
    AnswerFilePresent$ = "Y"
Else AnswerFilePresent$ = "N"
End If

Do
    _Limit 10
    CurrentTime$ = Date$ + "," + Left$(Time$, 5)
    Select Case Right$(CurrentTime$, 8)
        Case "23:59:58", "23:59:59"
            Midnight = 1
        Case Else
            Midnight = 0
    End Select
Loop While Midnight = 1

 Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -t" + CurrentTime$ + " -m -o -u2 -udfver102 -bootdata:2#p0,e,b" + CHR$(34) + DestinationFolder$_
 + "\ISO_Files\boot\etfsboot.com" + CHR$(34) + "#pEF,e,b" + CHR$(34) + DestinationFolder$ + "\ISO_Files\efi\microsoft\boot\efisys.bin"_
 + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34) + " " + CHR$(34) + FinalImageName$ + CHR$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

Select Case InjectionMode$
    Case "UPDATES"
        AddUpdatesStatusDisplay 0, 0, 15
    Case "DRIVERS"
        AddUpdatesStatusDisplay 0, 0, 24
    Case "BCD"
        AddUpdatesStatusDisplay 0, 0, 61
End Select

Print
Print "Performing a cleanup of files and mountpoints. Standby..."

' Perform a quick cleanup

_ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " by Hannes Sehestedt"

' Cleaning up previous stale DISM operations

Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Cleanup-WIM"
Shell _Hide Cmd$
Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /cleanup-mountpoints"
Shell _Hide Cmd$

' Cleanup folders

Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\Mount" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\Scratch" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\Temp" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\WIM_x64" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\Assets" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\WinPE" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\WinRE" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\WinRE_MOUNT" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\WinPE_MOUNT" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\Setup_DU_x64" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\SSU_x64" + Chr$(34) + " /s /q"
Shell _Hide Cmd$

' Remove the AV exclusion for the destination folder

CleanPath DestinationFolder$
Cmd$ = "powershell.exe -command Remove-MpPreference -ExclusionPath " + "'" + Chr$(34) + Temp$ + Chr$(34) + "'"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
If _FileExists("WIM_Exclude_Path.txt") Then Kill "WIM_Exclude_Path.txt"

' Clear the keyboard buffer

_KeyClear

' If the user opted to have a shutdown performed, then skip to the "ShutdownRequested" routine.
' Set an initial value of 0 for ShutdownStatus. It will be set to 1 if a shutdown is requested, 2 if a
' a hibernate is requested, or 3 if both are set. Having both set is considered an error condition and
' no shutdown or hibernation will take place.

ShutdownStatus = 0

If _FileExists(Environ$("userprofile") + "\Desktop\Auto_Shutdown.txt") Then ShutdownStatus = ShutdownStatus + 1
If _FileExists(Environ$("userprofile") + "\Desktop\Auto_Hibernate.txt") Then ShutdownStatus = ShutdownStatus + 2

Select Case ShutdownStatus
    Case 0, 3
        GoTo NoShutdown
    Case 1, 2
        GoTo ShutdownRequested
End Select

NoShutdown:

' Since any script playback is now completed, we will flip the status of the variable ScriptingChoice$ to "S".
' Normally, during playback of a script, we suppress pauses in the program with messages to the user. However,
' now that playback has been completed, we may have some messages that we want the user to see. To avoid simply
' skipping the pauses after these messages are displayed, we need to first change that variable so that the
' pause routine will actually stop and pause.

ScriptingChoice$ = "S"

' Clear any key presses from the buffer

_KeyClear

Do
    Cls
    Print
    Color 0, 10: Print "All processes have been completed.": Color 15

    ' Wait until users enters "X" (upper or lower case is fine) before we continue.

    Input "Enter X to Exit Back to the Main Menu: ", Temp$
Loop Until UCase$(Temp$) = "X"

If OpsPending$ = "Y" Then
    Cls
    Print
    Color 14, 4: Print "Warning!": Color 15
    Print
    Print "Please note that at least one Windows edition was detected that had pending operations during processing. Pending"
    Print "operations will prevent the necessary DISM Image Cleanup operations from being able to run during processing."
    Print "Typically, this is a result of adding NetFX3 to a Windows image."
    Print
    Print "A log file indicating any Windows editions that had pending operations can be found here:"
    Print
    Color 10: Print DestinationFolder$; "\logs\PendingOps.log": Color 15
    Print
    Print "It is suggested that you reapply your Windows updates to an image that does not already have NetFX3 enabled or"
    Print "has some other reason for pending operations. It is perfectly fine to enable NetFX3 after the DISM Image Cleanup"
    Print "operation has completed. Avoid using that output image as the source for applying further updates."
    Pause
End If

If AnswerFilePresent$ = "Y" Then
    Cls
    Print
    Color 14, 4: Print "CAUTION!";: Color 15: Print " Your final image file contains an autounattend.xml answer file. If this image is booted, depending upon the"
    Print "configuration, it is possible that it could wipe the boot drive and / or other drives in the system on which it is"
    Print "being booted."
    Print
    Print "You may want to name this image and mark any media created from it to avoid accidental usage."
    Pause
End If

ChDir ProgramStartDir$: GoTo BeginProgram

ShutdownRequested:

' User elected to have an automatic shutdown or hibernation performed. We will give then one last opportunity to abort.

Cls
Print
Color 0, 10: Print "All processes have been completed.": Color 15
Print

Select Case ShutdownStatus
    Case 1
        Print "A shutdown of the system has been requested. You can abort the shutdown if you delete or rename the file named"
        For x = 60 To 0 Step -1
            _Limit 1
            Locate 5, 1: Color 10: Print "AUTO_SHUTDOWN.txt";: Color 15: Print " located on your desktop within the next";: Color 10: Print x;: Color 15: Print "seconds.  "
            If Not (_FileExists(Environ$("userprofile") + "\Desktop\Auto_Shutdown.txt")) Then GoTo NoShutdown
        Next x
    Case 2
        Print "A hibernation of the system has been requested. You can abort the hibernation if you delete or rename the file named"
        For x = 60 To 0 Step -1
            _Limit 1
            Locate 5, 1: Color 10: Print "AUTO_HIBERNATE.txt";: Color 15: Print " located on your desktop within the next";: Color 10: Print x;: Color 15: Print "seconds.  "
            If Not (_FileExists(Environ$("userprofile") + "\Desktop\Auto_Hibernate.txt")) Then GoTo NoShutdown
        Next x
End Select

ff = FreeFile
Open ("WIM_Shutdown_log.txt") For Output As #ff
Print #ff, "Program run was completed on "; Date$; " at "; Time$

If OpsPending$ = "Y" Then
    Print #ff, ""
    Print #ff, "Warning!"
    Print #ff, ""
    Print #ff, "Please note that at least one Windows edition was detected that had pending operations during processing. Pending"
    Print #ff, "operations will prevent the necessary DISM Image Cleanup operations from being able to run during processing."
    Print #ff, "Typically, this is a result of adding NetFX3 to a Windows image."
    Print #ff, ""
    Print #ff, "A log file indicating any Windows editions that had pending operations can be found here:"
    Print #ff, ""
    Print #ff, DestinationFolder$; "\logs\PendingOps.log"
    Print #ff, ""
    Print #ff, "It is suggested that you reapply your Windows updates to an image that does not already have NetFX3 enabled or"
    Print #ff, "has some other reason for pending operations. It is perfectly fine to enable NetFX3 after the DISM Image Cleanup"
    Print #ff, "operation has completed. Avoid using that output image as the source for applying further updates."
End If

If AnswerFilePresent$ = "Y" Then
    Print #ff, ""
    Print #ff, "CAUTION! Your final image file contains an autounattend.xml answer file. If this image is booted, depending upon the"
    Print #ff, "configuration, it is possible that it could wipe the boot drive and / or other drives in the system on which it is"
    Print #ff, "being booted."
    Print #ff, ""
    Print #ff, "You may want to name this image and mark any media created from it to avoid accidental usage."
End If

Close #ff

Select Case ShutdownStatus
    Case 1
        Shell _Hide "shutdown /s /t 5 /f"
    Case 2
        Shell _Hide "shutdown /h"
End Select

GoTo EndProgram


' ***************************************************************************
' * Inject registry entries into BOOT.WIM to bypass Windows 11 requirements *
' ***************************************************************************

BypassWin11Requirements:

' Set the current folder to the location from which this program is run.

ChDir ProgramStartDir$

' Ask what options should be used

BootWimOptions:

Cls
Print "Select the option that you want by number:"
Print
Print "Note: For an explanation of these options, please see the related HELP topic."
Print
Print "1) Update the Boot.wim to bypass Win 11 system requirements"
Print "2) Force the use of the previous version of setup"
Print "3) Perform both of the above"
Print "4) Exit"
Print
Input "Make your selection by number:"; BootWimMods$

Select Case BootWimMods$
    Case "1", "2", "3"
        ' These are valid selections so we can jump out of this decision loop and continue.
        Exit Select
    Case "4"
        GoTo EndProgram
    Case Else
        Cls
        Print "Please enter a number from 1 to 4"
        Pause
        GoTo BootWimOptions
End Select

' Ask for source image file. Verify that it is a valid image.

GetImageName:

' Initialize variable
SourceImage$ = ""

Cls
Print "Please enter the full path and file name of the Windows ISO image to be updated."
Print
Line Input "Path and file name: ", SourceImage$
CleanPath SourceImage$
SourceImage$ = Temp$

FileName$ = Mid$(SourceImage$, (_InStrRev(SourceImage$, "\") + 1))

If Not (_FileExists(SourceImage$)) Then
    Cls
    Color 14, 4: Print "No such file exists.";: Color 15: Print " Please specify a valid file."
    Pause
    GoTo GetImageName
End If

SourcePath$ = Left$(SourceImage$, _InStrRev(SourceImage$, "\") - 1)
Cls
Print "Standby while we verify the validity of this file..."

' Start by determining the architectre type of the image (either a single or dual architecture type).

DetermineArchitecture SourceImage$, 1

Select Case ImageArchitecture$
    Case "x64", "x86"
        FileSourceType$ = "SINGLE"
    Case "DUAL", "NONE"
        Cls
        Color 14, 4: Print "The image specified is not valid.";: Color 15: Print " Please specify a valid image."
        Pause
        GoTo BypassWin11Requirements
End Select

MountISO SourceImage$

If Not (_FileExists(MountedImageDriveLetter$ + "\sources\install.wim")) Then
    Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourcePath$ + "'" + Chr$(34) + Chr$(34)
    Shell _Hide Cmd$
    Cls
    Color 14, 4: Print "The image specified is not valid.";: Color 15: Print " Please specify a valid image."
    Pause
    GoTo BypassWin11Requirements
End If

GetDestinationPath11:

Do
    Cls
    Print "Enter the path where the project will be created. This is where all the temporary files will be stored and we will"
    Print "save the final ISO image file here as well. NOTE: Use a folder that does not already exist or that contains no"
    Print "data that you need to keep as it may be deleted when we are done."
    Print
    Line Input "Enter the path where the project should be created: ", DestinationFolder$
Loop While DestinationFolder$ = ""

CleanPath DestinationFolder$
DestinationFolder$ = Temp$ + "\"

' We don't want user to specify the root of a drive

If Len(DestinationFolder$) = 3 Then
    Cls
    Color 14, 4
    Print "Please do not specify the root directory of a drive."
    Color 15
    Pause
    GoTo GetDestinationPath11
End If

' Check to see if the destination specified is on a removable disk

Cls
Print "Performing a check to see if the destination you specified is a removable disk."
Print
Print "Please standby..."
DriveLetter$ = Left$(DestinationFolder$, 2)
RemovableDiskCheck DriveLetter$
DestinationIsRemovable = IsRemovable

Select Case DestinationIsRemovable
    Case 2
        Cls
        Color 14, 4: Print "This is not a valid disk.";: Color 15: Print " Please specify another location."
        Pause
        GoTo GetDestinationPath11
    Case 1
        Cls
        Print "The disk that you specified is a removable disk. ";: Color 14, 4: Print "Please specify a fixed disk.": Color 15
        Print
        Print "NOTE: Project must be created on a fixed disk due to limitations of some Microsoft utilities."
        Pause
        GoTo GetDestinationPath11
    Case 0
        ' if the returned value was a 0, no action is necessary. The program will continue normally.
End Select

' Verify that the path specified exists.

If Not (_DirExists(DestinationFolder$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(DestinationFolder$)) Then
        Cls
        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
        Print
        Print "Please recheck the path you have specified and try again."
        Pause
        GoTo GetDestinationPath11
    End If
End If

' Ask user what they want to name the final ISO image file

Cls
UserSelectedImageName$ = "" ' Set initial value
Print "If you would like to specify a name for the final ISO image file that this project will create, please do so now,"
Print "WITHOUT an extension. You can also simply press ENTER to use the default name of Windows.ISO."
Print
Print "Enter name ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension, or press ENTER: ";: Line Input "", UserSelectedImageName$

If UserSelectedImageName = "" Then
    UserSelectedImageName$ = "Windows.iso"
Else
    UserSelectedImageName$ = UserSelectedImageName$ + ".iso"
End If

' Extract the resource files that are needed to set the "Installation Type" property to "Server"

Resource$ = _Embedded$("WimLib")
ff = FreeFile
Open "wimlib-imagex.exe" For Binary As #ff
Put #ff, 1, Resource$
Close #ff

Resource$ = _Embedded$("libwim15")
ff = FreeFile
Open "libwim-15.dll" For Binary As #ff
Put #ff, 1, Resource$
Close #ff

UsePreviousSetup:

If BootWimMods$ = "2" Or BootWimMods$ = "3" Then
    PreviousSetup$ = "Y"
Else
    PreviousSetup$ = "N"
End If

' Display Status Lines

Cls
Print "Setting an antivirus exclusion for the project"
Print "Performing file / folder cleanup and creating folders for project"
Print "Copying files from ISO image to project folder"
Print "Mounting and updating the BOOT.WIM"
Print "Unmounting the BOOT.WIM"
Print "Changing the Installation Type property to Server"
Print "Creating the final Windows image"
Print "Removing the antivirus exclusion and cleaning up temp files"

' If we have arrived here it means that the destination path already exists
' or we were able to create it successfully.

' Start by setting an AV exclusion for the destination path. We will log this location to a temporary file
' so that if the file is interrupted unexpectedly, we can remove the exclusion the next time the program
' is started.

' Add an AV exclusion for the destination folder

Locate 1, 1: Color 4: Print "Setting an antivirus exclusion for the project": Color 15
CleanPath DestinationFolder$
Cmd$ = "powershell.exe -command Add-MpPreference -ExclusionPath " + "'" + Chr$(34) + DestinationFolder$ + Chr$(34) + "'"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

' Remove the ISO_Files and Mount folders from the destination if they exist already and recreate them to make sure no old data
' is present in those locations.

Locate 1, 1: Color 10: Print "Setting an antivirus exclusion for the project": Color 15
Locate 2, 1: Color 4: Print "Performing file / folder cleanup and creating folders for project": Color 15
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "ISO_Files" + Chr$(34) + " /s /q > NUL"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "Mount" + Chr$(34) + " /s /q > NUL"
Shell _Hide Cmd$
Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "ISO_Files" + Chr$(34)
Shell _Hide Cmd$
Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "Mount" + Chr$(34)
Shell _Hide Cmd$
Locate 2, 1: Color 10: Print "Performing file / folder cleanup and creating folders for project": Color 15
Locate 3, 1: Color 4: Print "Copying files from ISO image to project folder": Color 15
Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + "\ " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "ISO_Files\ " + Chr$(34) + " /mir"
Shell _Hide Cmd$

' Unmount Windows image

Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourceImage$ + "'" + Chr$(34) + Chr$(34)
Shell _Hide Cmd$
Cmd$ = "attrib -h -s -r " + Chr$(34) + DestinationFolder$ + "ISO_Files\*.*" + Chr$(34) + " /S /D"
Shell _Hide Cmd$


' Update Setup, located in index #2 of the Boot.wim

Locate 3, 1: Color 10: Print "Copying files from ISO image to project folder": Color 15
Locate 4, 1: Color 4: Print "Mounting and updating the BOOT.WIM": Color 15
Cmd$ = "dism /mount-wim /wimfile:" + DestinationFolder$ + "ISO_Files\sources\boot.wim /index:2 /mountdir:" + DestinationFolder$ + "mount"
Shell _Hide Cmd$
Cmd$ = "reg load HKLM\offline " + DestinationFolder$ + "mount\windows\system32\config\system"
Shell _Hide Cmd$

If BootWimMods$ = "1" Or BootWimMods$ = "3" Then
    Cmd$ = "reg add HKLM\offline\Setup\LabConfig /v BypassTPMCheck /t reg_dword /d 0x00000001 /f"
    Shell _Hide Cmd$
    Cmd$ = "reg add HKLM\offline\Setup\LabConfig /v BypassSecureBootCheck /t reg_dword /d 0x00000001 /f"
    Shell _Hide Cmd$
    Cmd$ = "reg add HKLM\offline\Setup\LabConfig /v BypassRAMCheck /t reg_dword /d 0x00000001 /f"
    Shell _Hide Cmd$
End If

If PreviousSetup$ = "Y" Then
    Cmd$ = "reg add HKLM\offline\Setup /v CmdLine /t reg_sz /d " + Chr$(34) + "X:\Sources\Setup.exe" + Chr$(34) + " /f"
    Shell _Hide Cmd$
End If

Cmd$ = "reg unload HKLM\offline"
Shell _Hide Cmd$
Locate 4, 1: Color 10: Print "Mounting and updating the BOOT.WIM": Color 15
Locate 5, 1: Color 4: Print "Unmounting the BOOT.WIM": Color 15
Cmd$ = "dism /unmount-image /mountdir:" + Chr$(34) + DestinationFolder$ + "mount" + Chr$(34) + " /commit"
Shell _Hide Cmd$

' We will now update the image to allow for UPGRADE installation on unsupported hardware

' Here we use wimlib-imagex.exe to determine how many indices the install.wim holds. The output of the command is redirected to wimlib_results.txt
' and that file is then parsed to get the number of indices.

Locate 5, 1: Color 10: Print "Unmounting the BOOT.WIM": Color 15

If BootWimMods$ = "1" Or BootWimMods$ = "3" Then

    Locate 6, 1: Color 4: Print "Changing the Installation Type property to Server": Color 15
    Cmd$ = "wimlib-imagex.exe info " + Chr$(34) + DestinationFolder$ + "ISO_Files\sources\install.wim" + Chr$(34) + " --header | find " + Chr$(34) + "Image Count" + Chr$(34) + "" + " > wimlib_results.txt"
    Shell Cmd$
    ff = FreeFile
    Open "wimlib_results.txt" For Input As #ff
    Line Input #1, Temp$
    Close #ff

    ' Temp$ will hold the nuimber of indices as a string with no leading spaces.

    x = InStr(Temp$, "=")
    Temp$ = Right$(Temp$, (Len(Temp$) - x))
    Temp$ = LTrim$(Temp$)

    ' We are done with wimlib_results.txt. Delete it.

    If _FileExists("wimlib_results.txt") Then
        Kill "wimlib_results.txt"
    End If

    ' For each index, set the Installation Type prperty to server.

    For x = 1 To Val(Temp$)
        Cmd$ = "wimlib-imagex.exe info " + Chr$(34) + DestinationFolder$ + "ISO_Files\sources\install.wim" + Chr$(34) + Str$(x) + " --image-property WINDOWS/INSTALLATIONTYPE=server"
        Shell _Hide Cmd$
    Next x

    Locate 6, 1: Color 10: Print "Changing the Installation Type property to Server": Color 15

Else

    Locate 6, 1: Print "Changing the Installation Type property to Server (SKIPPED)"

End If

Locate 7, 1: Color 4: Print "Creating the final Windows image": Color 15
Cmd$ = Chr$(34) + OSCDIMGLocation$ + Chr$(34) + " -m -o -u2 -udfver102 -bootdata:2#p0,e,b" + Chr$(34) + DestinationFolder$_
 + "ISO_Files\boot\etfsboot.com" + Chr$(34) + "#pEF,e,b" + Chr$(34) + DestinationFolder$ + "ISO_Files\efi\microsoft\boot\efisys.bin"_
  + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "ISO_Files" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + UserSelectedImageName$ + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

' Remove the AV Exclusion

Locate 7, 1: Color 10: Print "Creating the final Windows image": Color 15
Locate 8, 1: Color 4: Print "Removing the antivirus exclusion and cleaning up temp files": Color 15
Cmd$ = "powershell.exe -command Remove-MpPreference -ExclusionPath " + "'" + Chr$(34) + DestinationFolder$ + Chr$(34) + "'"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
Locate 8, 1: Color 10: Print "Removing the antivirus exclusion and cleaning up temp files": Color 15

' Perform a cleanup by deleting the temporary ISO_Files and Mount folders as well as the resouce files that we extracted.
' These files are wimlib-imagex.exe and libwim-15.dll.

Cmd$ = "rd " + Chr$(34) + DestinationFolder$ + "ISO_Files" + Chr$(34) + " /S /Q > NUL"
Shell _Hide Cmd$
Cmd$ = "rd " + Chr$(34) + DestinationFolder$ + "Mount" + Chr$(34) + " /S /Q > NUL"
Shell _Hide Cmd$
If _FileExists("wimlib-imagex.exe") Then Kill "wimlib-imagex.exe"
If _FileExists("libwim-15.dll") Then Kill "libwim-15.dll"

ReplaceFile:

' Show the user where the file was saved and ask if they want to replace the original file with the updated file.

Cls
Print "The project has been completed. The final image is located here:"
Print
Color 10: Print DestinationFolder$; UserSelectedImageName$: Color 15
Print
Input "Do you want to replace the original file with this updated file"; Temp$
YesOrNo Temp$
Temp$ = YN$

If Temp$ = "X" Then GoTo ReplaceFile

If Temp$ = "Y" Then
    Cls
    Print "Moving the file ";: Color 10: Print DestinationFolder + "Windows.ISO": Color 15
    Print "   to"
    Color 10: Print SourceImage$: Color 15
    Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "Windows.ISO" + Chr$(34) + " " + Chr$(34) + SourceImage$ + Chr$(34)
    Shell _Hide Cmd$
End If

' That is all. All operations have finished.

Print
Print "All operations have been completed."
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' ***************************************************************************************
' * Make or update a bootable drive from one or more Windows / WinPE / WinRE ISO images *
' ***************************************************************************************

MakeBootDisk:

' This routine will allow you to create a bootable drive for installing Windows or to be used as an emergency
' boot disk. It can also create a boot disk that can hold many Windows images as well as WinPE and WinRE based
' images, allowing the user to select any one of those images for boot. The option to create one or more additional
' generic partitions is also provided as is support for BitLocker encryption on those additional partitions.

' IMPORTANT: This routine has some code to allow the user to choose whether any partitions aside from the
' first partition should be created as NTFS or exFAT. If you want the program to just automatically
' create all partitions as NTFS then set "UserCanPickFS$" below to FALSE. It you set it to TRUE then
' the user will be allowed to pick the filesystem to use.

UserCanPickFS$ = "FALSE" ' Set this to "FALSE" to always use NTFS. See the note above about this.
AddPart$ = ""
TotalPartitions = 0
AdditionalPartitions = 0
ReDim PartitionSize(4) As String
ReDim BitLockerFlag(4) As String
ReDim AutoUnlock(4) As String
ReDim Letter(4) As String
ReDim VolLabel(4) As String

VolLabel(1) = "UFD VOL 1"
VolLabel(2) = "UFD VOL 2"
VolLabel(3) = "UFD VOL 3"
VolLabel(4) = "UFD VOL 4"

SingleOrMulti:

SingleOrMulti$ = "" ' Set initial value

Cls
Print "Do you want to create a boot disk that boots a single Windows ISO image, or do you want the choice to boot from"
Print "multiple different ISO images which can include WinPE or WinRE based media?"
Print
Print "Choose one:"
Print
Color 0, 10: Print "1)";: Color 15: Print " ";: Color 0, 14: Print "Create or refresh standard boot media created from a single Windows ISO image       "
Color 0, 10: Print "2)";: Color 15: Print " ";: Color 0, 14: Print "Create or refresh boot media to allow booting from an unlimited number of ISO images": Color 15
Print
Input "Enter your selection by number: ", SingleOrMulti$

Select Case SingleOrMulti$
    Case "1"
        SingleOrMulti$ = "SINGLE"
        GoTo CreateSingleImageDisk
    Case "2"
        SingleOrMulti$ = "MULTI"
        GoTo WipeOrRefresh
    Case Else
        Print
        Color 14, 4: Print "Invalid selection!": Color 15
        Print "Your selection was not valid. Please provide a valid response."
        Pause
        GoTo SingleOrMulti
End Select

CreateSingleImageDisk:

AutounattendHandling
EiCfgHandling

' Get Windows ISO path to copy to the thumb drive

GetSourceISOForMakeBoot:

MakeBootableSourceISO$ = "" ' Set initial value

Do
    Cls
    Print "Enter the full path including the file name for the Windows ISO image you want to copy to the drive."
    Print
    Line Input "Enter the full path: ", MakeBootableSourceISO$
Loop While MakeBootableSourceISO$ = ""

CleanPath MakeBootableSourceISO$
MakeBootableSourceISO$ = Temp$

If Not _FileExists(MakeBootableSourceISO$) Then
    Cls
    Print "We could not find the file that you specified. Please check the path and filename and try again."
    Pause
    GoTo GetSourceISOForMakeBoot
End If

' If we reach this point, then the path provided is valid and the file name specified exists.

' Perform a check to see if the image file specified is valid and if it is a single or dual architecture image.
' If the images contains a \sources\install.wim file, then the image is a single architecture image (either x86 or x64).
' If the image contains both an \x86\sources\install.wim AND a \x64\sources\install.wim then it is a dual architecture
' image. If neither is true (for example if the image contains ESD files rather than WIM files) then we will consider
' the image to be invalid.

Cls
Print "Standby for a moment while we check to make sure that the file specified is valid."
MountISO MakeBootableSourceISO$
CDROM$ = MountedImageDriveLetter$

If _FileExists(CDROM$ + "\sources\install.wim") Then
    Architecture = 1
ElseIf (_FileExists(CDROM$ + "\x86\sources\install.wim")) And (_FileExists(CDROM$ + "\x64\sources\install.wim")) Then
    Architecture = 2
Else
    Architecture = 0
End If

' The next command dismounts the ISO image since we are now done checking to see if it is a valid image.

Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + MakeBootableSourceISO$ + "'" + Chr$(34) + Chr$(34)
Shell _Hide Cmd$

If Architecture = 0 Then
    Cls
    Color 14, 4: Print "The image that you have specified is not valid.";: Color 15: Print " It does not contain either"
    Print "a single architecture (x86 or x64) or dual architecture (both x86 and x64)"
    Print "configuration. Note that we are expecting that the image contains an install.wim"
    Print "file(s). This program is NOT designed to work with images that use install.esd files."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

WipeOrRefresh:

' If we reach this point, then the image specified by the user is valid.

' The user will now need to make a choice. They can configure a drive from scratch, creating all partitions and copying the
' files from an image file to that thumb drive, or they can simply refresh an existing Windows installation. Selecting the option
' to refresh an installation leave all the partitions alone but it will erase the contents of the FAT32 and exFAT Windows
' installation partitions and copy a new set of files to those partitions from the image chosen by the user.

WipeOrRefresh = 0

Do
    Cls
    Print "You have two options:"
    Print
    Color 0, 10: Print "1)";: Color 15: Print " ";: Color 0, 14: Print "WIPE DISK:";: Color 15
    Print " This will completely wipe the contents of a disk and configure it from scratch. Use this the first time"
    Print "   you are preparing a disk with this program."
    Print
    Color 0, 10: Print "2)";: Color 15: Print " ";: Color 0, 14: Print "REFRESH DISK:";: Color 15
    Print " If you are booting a single Windows image, this option will allow you to replace that image with an"
    Print "   updated image or a completely different image while leaving data on all other partitions alone. If you have chosen"
    Print "   the option to create a disk allowing you to select from multiple images, then a refresh will update the customized"
    Print "   Windows PE installation that makes this possible but it will leave everything else alone. Refresh is intended only"
    Print "   for disks previously initialized by a WIPE operation using this program."
    Print
    Input "Which option do you want (1 or 2)"; WipeOrRefresh
Loop Until (WipeOrRefresh = 1 Or WipeOrRefresh = 2)

SelectFS:

FSType$ = "NTFS" ' Set default

If WipeOrRefresh = 1 And UserCanPickFS$ = "TRUE" Then
    FSType$ = ""
    Do
        Cls
        Print "For partitions other than the first partition, do you want to use the exFAT or NTFS filesystem?"
        Print
        Input "Enter exFAT, NTFS, or HELP for suggestions: "; FSType$
        FSType$ = UCase$(FSType$)
    Loop Until FSType$ = "EXFAT" Or FSType$ = "NTFS" Or FSType$ = "HELP"
    If FSType$ = "HELP" Then
        Cls
        Print "For an SSD or HDD, it is suggested that you select NTFS. For flash drives, you may prefer exFAT."
        Print
        Print "NOTE: You may need to experiment. NTFS should work in just about every case. The ability to format as exFAT was"
        Print "added mainly for experimental issues. It's possible that the performance of some flash media may degrade when"
        Print "NTFS is used because the garbage collection mechanisms are designed for exFAT rather than NTFS. If you have any"
        Print "difficulties using exFAT then I would suggest re-initializing the device using NTFS."
        Pause
        GoTo SelectFS
    End If
End If

If WipeOrRefresh = 1 GoTo SelectADisk

' We arrive here if the user elected to perform a refresh

GetDriveInfo_FAT32:

' Since we are performing a refresh operation, the drive letters for the Windows installation media should already be present.
' We will now search for these drive letters. In addition, we need to make sure that there is not more than one Windows
' installation disk connected to the system to make sure that we do not accidentally select the wrong disk.

' Initialize variables

Par1SingleInstancesFound = 0
Par2SingleInstancesFound = 0
Par1MultiInstancesFound = 0
Par2MultiInstancesFound = 0

Restore DriveLetterData

For x = 1 To 24
    Read MediaLetter$
    If _FileExists(MediaLetter$ + ":\VOL1_S_MEDIA.WIM") Then
        Par1SingleInstancesFound = Par1SingleInstancesFound + 1
    ElseIf _FileExists(MediaLetter$ + ":\VOL1_M_MEDIA.WIM") Then
        Par1MultiInstancesFound = Par1MultiInstancesFound + 1
    End If
Next x

Restore DriveLetterData

For x = 1 To 24
    Read MediaLetter$
    If _FileExists(MediaLetter$ + ":\VOL2_S_MEDIA.WIM") Then
        If _DirExists(MediaLetter$ + ":\PE_BACKUP") GoTo CannotRefresh
        Par2SingleInstancesFound = Par2SingleInstancesFound + 1
    ElseIf _FileExists(MediaLetter$ + ":\VOL2_M_MEDIA.WIM") Then
        If _DirExists(MediaLetter$ + ":\PE_BACKUP") GoTo CannotRefresh
        Par2MultiInstancesFound = Par2MultiInstancesFound + 1
    End If
Next x

' Check if ONLY the partitions for a single image type are found

If Par1SingleInstancesFound = 1 And Par2SingleInstancesFound = 1 Then
    If Par1MultiInstancesFound = 0 And Par2MultiInstancesFound = 0 Then
        GoTo DetectBootMediaLetters
    End If
End If

' Check if ONLY the partitions for a multi image type are found

If Par1MultiInstancesFound = 1 And Par2MultiInstancesFound = 1 Then
    If Par1SingleInstancesFound = 0 And Par2SingleInstancesFound = 0 Then
        GoTo DetectBootMediaLetters
    End If
End If

' We arrive here if an illegal state is reached (there are partitions for both single and multi image types
' or we don't have both partitions for either).

CannotRefresh:

Cls
Print "You have chosen to refresh a disk but we have encountered one of the following problems:"
Print
Print " 1) We were not able to find any previously created disk to refresh."
Print " 2) We found more than one disk previously created by this program which prevents us from determining which to use."
Print " 3) The disk is configured to boot an ISO image, it must be in a state where no image has been selected for boot."
Print
Print "Please correct this situation and then run this batch file again."
Print
Print "Please note that for condition #3 above, you can simply run the 'Config_UFD' batch file on your disk to revert the disk"
Print "back to the original state. Then run the REFRESH operation here once again."
Print
Print "We will now return you to the main menu."
Pause
ChDir ProgramStartDir$: GoTo BeginProgram

DetectBootMediaLetters:

' If we get here, then we already know that there is boot media previously created by this program connected to the system.
' We also know that there is only one such media connected. We will scan all the drives to determine what the drive letter
' for each of the first two partitions are.

Restore DriveLetterData

If SingleOrMulti$ = "SINGLE" Then

    For x = 1 To 24
        Read MediaLetter$
        If _FileExists(MediaLetter$ + ":\VOL1_S_MEDIA.WIM") Then
            FAT32DriveLetter$ = MediaLetter$
            GoTo CheckVol2_S
        End If
    Next x

    Cls
    Print "You have elected to perform a refresh operation of a boot disk type for which we found no disk."
    Print "Please ensure that you select the option (SINGLE or MULTI image type) that matches your disk."
    Print
    Print "You will be returned to the main menu."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram

    CheckVol2_S:

    Restore DriveLetterData

    For x = 1 To 24
        Read MediaLetter$
        If _FileExists(MediaLetter$ + ":\VOL2_S_MEDIA.WIM") Then
            exFATorNTFSdriveletter$ = MediaLetter$
            GoTo Continue_Refresh
        End If
    Next x

    Cls
    Print "You have elected to perform a refresh operation of a boot disk type for which we found no disk."
    Print "Please ensure that you select the option (SINGLE or MULTI image type) that matches your disk."
    Print
    Print "You will be returned to the main menu."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If


If SingleOrMulti$ = "MULTI" Then

    For x = 1 To 24
        Read MediaLetter$
        If _FileExists(MediaLetter$ + ":\VOL1_M_MEDIA.WIM") Then
            FAT32DriveLetter$ = MediaLetter$
            GoTo CheckVol2_M
        End If
    Next x

    Cls
    Print "You have elected to perform a refresh operation of a boot disk type for which we found no disk."
    Print "Please ensure that you select the option (SINGLE or MULTI image type) that matches your disk."
    Print
    Print "You will be returned to the main menu."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram

    Restore DriveLetterData

    CheckVol2_M:

    For x = 1 To 24
        Read MediaLetter$
        If _FileExists(MediaLetter$ + ":\VOL2_M_MEDIA.WIM") Then
            exFATorNTFSdriveletter$ = MediaLetter$
            GoTo Continue_Refresh
        End If
    Next x

    Cls
    Print "You have elected to perform a refresh operation of a boot disk type for which we found no disk."
    Print "Please ensure that you select the option (SINGLE or MULTI image type) that matches your disk."
    Print
    Print "You will be returned to the main menu."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram

End If

Continue_Refresh:

' When we arrive here, then we have the drive letters for the disk to be refreshed. We will now mount the
' ISO image so that we can refresh the drive.

' We already know the architecture of the ISO image being used for the refresh from the earlier mount of the image
' so there is no need to check it again. Architecture = 1 if it is a single architecture image, 2 if dual architecture
' We will now clean off the thumb drive partitions and copy the refreshed data to it. This gives us everything we need
' to jump to the already existing existing routine that mounts the ISO Image and copies file to the drive EXCEPT
' that this routine is expecting the drive letter of the FAT32 partition in Letter$(1) and the exFAT / NTFS partition
' in Letter$(2).

' We will also format those partitions first to clear the current content

' We want to keep the exiting volume labels, but these are destroyed when we format the partitions. We will save the current volume labels
' and then restore them after the format.

Cls
Color 0, 10: Print "Disk to be refreshed has been automatically located.": Color 15
Print
Shell "vol " + FAT32DriveLetter$ + ": > temp.txt"
Open "temp.txt" For Input As #1
FileLength = LOF(1)
VolLabel$(1) = Input$(FileLength, 1)
Close #1
Kill "temp.txt"

If InStr(VolLabel$(1), "has no label") Then
    VolLabel$(1) = ""
Else
    VolLabel$(1) = Mid$(VolLabel$(1), 23, ((InStr(VolLabel$(1), Chr$(13))) - 23))
End If

Shell "vol " + exFATorNTFSdriveletter$ + ": > temp.txt"
Open "temp.txt" For Input As #1
FileLength = LOF(1)
VolLabel$(2) = Input$(FileLength, 1)
Close #1
Kill "temp.txt"

If InStr(VolLabel$(2), "has no label") Then
    VolLabel$(2) = ""
Else
    VolLabel$(2) = Mid$(VolLabel$(2), 23, ((InStr(VolLabel$(2), Chr$(13))) - 23))
End If

' Format partition 1

Cmd$ = "format " + FAT32DriveLetter$ + ": /FS:FAT32 /Q /Y > NUL"
Shell Cmd$

' Restore the volume label if the original volume label was not blank.

If VolLabel$(1) <> "" Then
    Cmd$ = "label " + FAT32DriveLetter$ + ": " + VolLabel$(1)
    Shell Cmd$
End If

If SingleOrMulti$ = "SINGLE" Then
    Cmd$ = "format " + exFATorNTFSdriveletter$ + ": /FS:" + FSType$ + " /Q /Y > NUL"
    Shell Cmd$

    ' Restore the volume label if the original volume label was not blank.

    If VolLabel$(2) <> "" Then
        Cmd$ = "label " + exFATorNTFSdriveletter$ + ": " + VolLabel$(2)
        Shell Cmd$
    End If
End If

ReDim Letter(2) As String

Letter$(1) = FAT32DriveLetter$
Letter$(2) = exFATorNTFSdriveletter$

GoTo DoneWithBitLocker

SelectADisk:

' If the user picks the option to wipe the drive and set it up from scratch, then we come here.

GoSub SelectDisk

GoSub CheckTotalDiskSize

' We are done with the variable ListOfDisks$. Let's free up the space it used by clearing it
' becasue this variable could potentially contain a good amount of text.

ListOfDisks$ = ""

AskForAdditionalPartitions:

AddPart$ = "" ' Set initial value

Cls
Print "We will create 2 partitions to facilitate making a boot disk that can be booted on both BIOS and UEFI based systems and"
Print "that supports both x64 and x86 editions of Windows. If you wish, additional partitions can be created to store other"
Print "data. Please note that this program supports a maximum of 2 additional partitions, for a total of 4 partitions."
Print
Input "Do you want to create additional partitions"; AddPart$

' Parse the users response to determine if it is a valid yes / no response.

YesOrNo AddPart$
AddPart$ = YN$

If AddPart$ = "X" Then
    Print
    Color 14, 4: Print "Please provide a valid response.": Color 15
    Pause
    GoTo AskForAdditionalPartitions
End If

' The user entered a valid response

If AddPart$ = "N" Then
    ' User does not want to create additional partitions so move on to asking what disk thay want to use.
    ' Also, we know that the total number of partitions will be 2.
    TotalPartitions = 2
    GoTo PartitionSizes
End If

' The user wants to add partitions. Ask how many additional partitions.

HowManyPartitions:

AdditionalPartitions = 0 ' Set initial value
Cls
Input "How many additional partitions do you want to create (2 maximum)"; AdditionalPartitions

If (AdditionalPartitions < 1) Or (AdditionalPartitions > 2) Then
    Print
    Color 14, 4: Print "The number of additional partitions must be 1 or 2.": Color 15
    Pause
    GoTo HowManyPartitions
End If

' For each partition, except the last partition, ask for the size. The last partition will be assigned
' all remaining space.

Cls
Print "On the next screen you will be asked for partition sizes."
Print
Color 0, 10: Print "IMPORTANT:";: Color 15: Print " For the first partition, it is suggested to use a size larger than needed. A suggested size is 2.5 GB. This"
Print "will allow you to include customized and sysprep images that may need more space than standard Windows images. If you"
Print "make this partition too small, and then later want to use an image(s) that needs more space, you will not be able to use"
Print "the refresh feture, forcing you to recreate the media from scratch."
Print
Print "You should make the second partition large enough to hold your Windows image(s). If you are creating a disk that allows"
Print "you to select from multiple images, then this partition needs to have enough room to store ALL of your images, PLUS the"
Print "size of your single largest image because we will store an extracted copy of that image here."
Print
Print "TIP: You may want to make your second partition a bit larger than you currently need in case any of your images get"
Print "larger in the future or if you want to add additional images."
Pause

PartitionSizes:

TotalPartitions = AdditionalPartitions + 2

' Get partition sizes. We need to remove the leading space for the partition size values so we are going to convert it
' to a string. The last partition will be set to occupy all remaining space on the drive. If only two partitions are
' being created then we don't need to ask about encryption.

For x = 1 To (TotalPartitions - 1)

    RedoPartitionSize:

    GoSub ShowPartitionSizes
    Print
    Print "Enter the size below followed by "; Chr$(34); "M"; Chr$(34); " for Megabytes, "; Chr$(34); "G"; Chr$(34); " for Gigabytes, or "; Chr$(34); "T"; Chr$(34); " for Terabytes."
    Print
    Print "Examples: 500M, 20G, 1T, 700m, 1g"
    Print
    Print "Enter the size of partition number";: Color 0, 10: Print x;: Color 15: Print ": ";
    Input "", TempPartitionSize$
    TempUnit$ = UCase$(Right$(TempPartitionSize$, 1))
    TempValue = Val(TempPartitionSize$)
    Select Case TempUnit$
        Case "M"
            PartitionSize$(x) = Str$(TempValue)
            GoTo PartitionUnitsValid
        Case "G"
            PartitionSize$(x) = Str$(TempValue * 1024)
            GoTo PartitionUnitsValid
        Case "T"
            PartitionSize$(x) = Str$(TempValue * 1048576)
            GoTo PartitionUnitsValid
        Case Else
            Cls
            Print "Enter a valid size including an "; Chr$(34); "M"; Chr$(34); " for Megabytes, "; Chr$(34); "G"; Chr$(34); " for Gigabytes, or "; Chr$(34); "T"; " for Terabytes."
            Pause
            GoTo RedoPartitionSize
    End Select

    PartitionUnitsValid:

    If (Val(PartitionSize$(x))) <= 0 Then
        Color 14, 4: Print "You must enter a size larger than 0.": Color 15
        Print
        GoTo RedoPartitionSize
    End If
Next x

' For each added partition, ask if it should be BitLocker encrypted.

AskAboutEncryption:

For x = 1 To AdditionalPartitions

    RedoAskAboutEncryption:

    GoSub ShowPartitionSizes
    Print
    Print "You have specified that"; AdditionalPartitions; "additional partitions should be added."
    Print
    Print "Do you want to BitLocker encrypt partition #";: Color 0, 10: Print x + 2;: Color 15
    Input BitLockerFlag$(x + 2)
    If BitLockerFlag$(x + 2) = "" Then GoTo RedoAskAboutEncryption
    BitLockerFlag$(x + 2) = UCase$(BitLockerFlag$(x + 2))
    BitLockerFlag$(x + 2) = Left$(BitLockerFlag$(x + 2), 1)
    Select Case BitLockerFlag$(x + 2)
        Case "Y"
            GoTo ResponseIsValid2
        Case "N"
            ' User does not want to BitLocker encrypt this partitition.
        Case Else
            Print
            Color 14, 4: Print "Please provide a valid response.": Color 15
            Pause
            Cls
            GoTo RedoAskAboutEncryption
    End Select

    ResponseIsValid2:

    Cls
    If BitLockerFlag$(x + 2) = "Y" Then
        AskAboutAutoUnlock:
        Color 0, 10: Print "BitLocker Drive Encryption Selected": Color 15
        Print
        Input "Do you also want to autounlock this drive on this system"; AutoUnlock$(x + 2)
        If AutoUnlock$(x + 2) = "" Then GoTo ResponseIsValid2
        AutoUnlock$(x + 2) = UCase$(AutoUnlock$(x + 2))
        AutoUnlock$(x + 2) = Left$(AutoUnlock$(x + 2), 1)
        Select Case AutoUnlock$(x + 2)
            Case "Y"
                GoTo ResponseIsValid3
            Case "N"
                ' User does not want to autounlock the partition.
            Case Else
                Print
                Color 14, 4: Print "Please provide a valid response.": Color 15
                Pause
                Cls
                GoTo AskAboutAutoUnlock
        End Select

        ResponseIsValid3:

    End If
Next x

AfterBitLockerInfo:

' Init variable

TotalSpaceNeeded = Val(PartitionSize$(1))

If TotalPartitions > 2 Then

    For x = 2 To (TotalPartitions - 1)
        TotalSpaceNeeded = TotalSpaceNeeded + Val(PartitionSize$(x))
    Next x

End If

If TotalSpaceNeeded > AvailableSpace Then
    Cls
    Color 14, 4: Print "Warning!";: Color 15: Print " You have have specified partition sizes that total more than the space available on the selected disk."
    Print
    Print "Please check the values that you have supplied and the disk that you selected and try again."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

Removable:

' Write the commands needed to initialize the disk to the file named "TEMP.BAT"

' NOTE: There is a problem where performing a "clean" in diskpart will sometimes fail the first time.
' It usually works the second time but in testing with batch files we have seen failures even on the
' second time. The failures only seem to happen on MBR disks. As a result, we are performing a clean
' operation twice.

Cls
Print "Initializing disk..."
Open "TEMP.BAT" For Output As #1
Print #1, "@echo off"
Print #1, "(echo select disk"; DiskID
Print #1, "echo clean"
Print #1, "echo clean"

If Override$ = "Y" Then
    Print #1, "echo convert gpt"
Else
    Print #1, "echo convert mbr"
End If

Print #1, "echo exit"
Print #1, "echo ) | diskpart > NUL"
Close #1
Shell "TEMP.BAT"
If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

GoSub SelectAutoOrManual

' We reach this point when all drive letters have been successfully assigned.

' Diskpart wants whole numbers only for disk size values so we are stripping off any non-integer portion

For x = 1 To TotalPartitions
    TempLong = Val(PartitionSize$(x))
    TempLong = Int(TempLong)
    PartitionSize$(x) = Str$(TempLong)
Next x

' Ask the user for volume label names.

Cls
Print "Total number of volumes to prepare:";: Color 0, 10: Print TotalPartitions: Color 15
Print
Print "For each volume, please provide a label to be assigned to the volume. To accept the default name that is shown,"
Print "simply press enter."
Print

For x = 1 To TotalPartitions

    GetLabel:

    Print ""
    Locate x + 5, 1

    Print "Enter volume label for partition number";: Color 0, 10: Print x;: Color 15
    If x = 1 Then
        Print " - 11 characters maximum. ";
    ElseIf FSType$ = "NTFS" Then
        Print " - 32 characters maximum. ";
    ElseIf FSType$ = "EXFAT" Then
        Print " - 11 characters maximum. ";
    End If
    Print "("; VolLabel$(x); ") : ";
    Input "", NewLabel$

    ' If the user hit ENTER to accept the default volume label, then we need to set the new volume label
    ' to the value of the default volume label.

    If NewLabel$ = "" Then
        NewLabel = VolLabel$(x)
    End If

    ' If x=1 then we are working on the first volume label which is limited to 11 characters. Anything  after 1 will be either an exFAT partition
    ' which also has an 11 character limit, or NTFS which is limited to 32 characters. We are using the variables "Row" "RowEnd and "Column" to
    ' position the cursor on the screen. We do this because if the user enters an invalid value, we want erase the invalid response from the
    ' screen and move the prompt back to the same place on the screen.

    Select Case x
        Case 1
            If Len(NewLabel$) > 11 Then
                Row = x + 5
                RowEnd = CsrLin
                Do While Row <= RowEnd
                    Print
                    Column = 1
                    Do While Column < 121
                        Locate Row, Column: Print " ";
                        Column = Column + 1
                    Loop
                    Row = Row + 1
                Loop
                GoTo GetLabel
            End If
        Case Else
            If FSType$ = "EXFAT" Then
                MaxLabelLength = 11
            Else
                MaxLabelLength = 32
            End If
            If Len(NewLabel$) > MaxLabelLength Then
                Row = x + 5
                RowEnd = CsrLin
                Do While Row <= RowEnd
                    Print
                    Column = 1
                    Do While Column < 121
                        Locate Row, Column: Print " ";
                        Column = Column + 1
                    Loop
                    Row = Row + 1
                Loop
                GoTo GetLabel
            Else

            End If
    End Select

    ' Since the first volume is FAT32, note that FAT32 volumes only accept uppercase volume labels. The "label" command automatically converts any text to
    ' uppercase, but to avoid comparison errors if try to compare a volume label entered by a user with the actual volume label, we will store the name in
    ' uppercase in the variable.

    If x = 1 Then
        NewLabel$ = UCase$(NewLabel$)
    End If

    If NewLabel$ <> "" Then
        VolLabel(x) = NewLabel$
    Else
        VolLabel(x) = ""
    End If

    ValidVolLabel:

Next x

Cls
Color 0, 10: Print "Performing initial preparation of the disk. Note that this may take a while.": Color 15
Print
Open "TEMP.BAT" For Output As #1
Print #1, "@echo off"
Print #1, "(echo select disk"; DiskID

For x = 1 To TotalPartitions
    If x < TotalPartitions Then
        Print #1, "echo create partition primary size="; PartitionSize$(x)
        If x = 1 Then
            Print #1, "echo active"
            If VolLabel$(x) = "" Then
                Print #1, "echo format fs=fat32 quick override"
            Else
                Print #1, "echo format fs=fat32 quick override label=" + Chr$(34) + VolLabel$(x) + Chr$(34)
            End If
            Print #1, "echo assign letter="; Letter$(x)
        Else
            If VolLabel$(x) = "" Then
                Print #1, "echo format FS="; FSType$; " quick override"
            Else
                Print #1, "echo format FS="; FSType$; " quick override label=" + Chr$(34) + VolLabel$(x) + Chr$(34)
            End If
            Print #1, "echo assign letter="; Letter$(x)
        End If
    Else
        Print #1, "echo create partition primary"
        If VolLabel$(x) = "" Then
            Print #1, "echo format FS="; FSType$; " quick override"
        Else
            Print #1, "echo format FS="; FSType$; " quick override label=" + Chr$(34) + VolLabel$(x) + Chr$(34)
        End If
        Print #1, "echo assign letter="; Letter$(x)
    End If
Next x

Print #1, "echo exit"
Print #1, "echo ) | diskpart > NUL"
Close #1
Shell "TEMP.BAT"
If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

' Handle BitLocker encryption

If TotalPartitions = 2 Then GoTo DoneWithBitLocker

' We know that there are additional partitions. Determine how many are to be encrypted.

BitLockerCount = 0 ' Set initial value

For x = 3 To TotalPartitions
    If BitLockerFlag$(x) = "Y" Then
        BitLockerCount = BitLockerCount + 1
    End If
Next x

If BitLockerCount = 0 Then GoTo DoneWithBitLocker

For x = 3 To TotalPartitions
    If BitLockerFlag$(x) = "Y" Then
        Open "TEMP.BAT" For Output As #1
        Print #1, "@echo off"
        Print #1, ""
        Print #1, "set /a counter=0"
        Print #1, ""
        Print #1, ":StartEncryption"
        Print #1, "%SystemRoot%\system32\manage-bde -on "; Letter$(x); ": -pw -used -em xts_aes128"
        Print #1, ""
        Print #1, "echo."
        Print #1, "echo ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #1, "echo :: We will wait a maximum of 5 minutes for BitLocker to initialize. ::"
        Print #1, "echo :: Typically, BitLocker will initialize much quicker, but we are    ::"
        Print #1, "echo :: allowing time for very slow media on slow systems to initialize. ::"
        Print #1, "echo ::                                                                  ::"
        Print #1, "echo :: IMPORTANT: If an error was displayed above, we still need to     ::"
        Print #1, "echo :: wait for BitLocker to timeout. Take note of any error message    ::"
        Print #1, "echo :: now to help you to take corrective action.                       ::"
        Print #1, "echo ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #1, "echo."
        Print #1, ""
        Print #1, "for /L %%a in (1,1,150) do ("
        Print #1, "timeout /t 2 /nobreak > nul"
        Print #1, "manage-bde -status "; Letter$(x); ": -p > nul"
        Print #1, "if errorlevel 0 goto :success"
        Print #1, ")"
        Print #1, ""
        Print #1, "echo."
        Print #1, "echo ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #1, "echo :: BitLocker was not enabled. Please check above for any error messages that may indicate ::"
        Print #1, "echo :: why it failed. When ready to proceed, press any key. Note that we will try to enable   ::"
        Print #1, "echo :: BitLocker 3 times. If it still fails, the program will continue, but you will need to  ::"
        Print #1, "echo :: enable BitLocker encryption later.                                                     ::"
        Print #1, "echo ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #1, "echo."
        Print #1, "set /p b="; Chr$(34); "Press any key to continue..."; Chr$(34); ""
        Print #1, "manage-bde -off "; Letter$(x); ": > nul"
        Print #1, "set /a counter=counter+1"
        Print #1, "if %counter% LSS 3 goto StartEncryption"
        Print #1, "cls"
        Print #1, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #1, "echo :: BitLocker failed to encrypt drive "; Letter$(x); ":. The program will continue, but please be aware that ::"
        Print #1, "echo :: if you want BitLocker encryption enabled on this drive, you will need to enable it later. ::"
        Print #1, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #1, "echo."
        Print #1, "set /p b="; Chr$(34); "Press any key to continue..."; Chr$(34); ""
        Print #1, "goto done"
        Print #1, ""
        Print #1, ":success"
        If AutoUnlock$(x) = "Y" Then
            Print #1, "%SystemRoot%\system32\manage-bde.exe -autounlock -enable "; Letter$(x); ":"
        End If
        Print #1, ""
        Print #1, ":done"
        Print #1, "echo."
        Print #1, "echo BitLocker drive encryption for partition"; x; "completed."
        Print #1, "echo."
        Close #1
        Shell "TEMP.BAT"
        If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"
    End If
Next x

DoneWithBitLocker:

' If the user wants to create a multi image media, jump to that code now.

If SingleOrMulti$ = "MULTI" Then GoTo CreateMultiImageDisk

' Creating single image media.

Select Case Architecture
    Case 2
        Print "Creating a dual architecture boot"
    Case 1
        Print "Creating a single architecture boot"
End Select

MountISO MakeBootableSourceISO$
CDROM$ = MountedImageDriveLetter$

' If Architecture = 2 then we need to skip to the section to process a dual architecture image properly

If Architecture = 2 Then GoTo DualArchitecture

ff = FreeFile
Open "TEMP.BAT" For Output As #ff
Print #ff, "@echo off"
Print #ff, "del WIM_File_Copy_Error_1.txt > NUL 2>&1"
Print #ff, "del WIM_File_Copy_Error_2.txt > NUL 2>&1"
Print #ff, "echo."
Print #ff, "echo *************************************************************"
Print #ff, "echo * Copying files. Be aware that this can take quite a while, *"
Print #ff, "echo * especially on the 2nd partition and with slower media.    *"
Print #ff, "echo * Please be patient and allow this process to finish.       *"
Print #ff, "echo *************************************************************"
Print #ff, "echo."
Print #ff, "echo Copying files to partition #1"
Print #ff, "set CurrentPartition=1"

If ExcludeAutounattend$ = "N" Then
    Print #ff, "robocopy "; CDROM$; "\ "; Letter$(1); ":\ /mir /xd sources /njs /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Else
    Print #ff, "robocopy "; CDROM$; "\ "; Letter$(1); ":\ /mir /xf autounattend.xml /xd sources /njs /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
End If

Print #ff, "robocopy "; CDROM$; "\sources "; Letter$(1); ":\sources boot.wim /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "echo Copying files to partition #2"
Print #ff, "set CurrentPartition=2"
Print #ff, "robocopy "; CDROM$; "\sources "; Letter$(2); ":\sources /mir /njh /njs /xf boot.wim /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\boot "; Letter$(2); ":\boot /mir /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\efi "; Letter$(2); ":\efi /mir /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\support "; Letter$(2); ":\support /mir /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "echo. > "; Letter$(1); ":\VOL1_S_MEDIA.WIM"
Print #ff, "echo. > "; Letter$(2); ":\VOL2_S_MEDIA.WIM"
Print #ff, "goto cleanup"
Print #ff, ":HandleError"
Print #ff, "echo An error occurred > WIM_File_Copy_Error_%CurrentPartition%.txt"
Print #ff, ":cleanup"
Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'"; MakeBootableSourceISO$; "'"; Chr$(34); Chr$(34) + " > NUL"

' Commenting out the section below that patches for Secure Boot. It is no longer needed with the ADK and Windows
' PE release of May, 2024
'
'Print #ff, "cls"
'Print #ff, "REM Start of routine to patch for Secure Boot mitigations"
'Print #ff, ""
'Print #ff, "echo *************************************************************************************"
'Print #ff, "echo *                              Applying Secure Boot Fixes                           *"
'Print #ff, "echo *                                                                                   *"
'Print #ff, "echo * Checking to see if this system has BlackLotus UEFI Bootkit mitigations installed. *"
'Print #ff, "echo * If it does, we will use this to patch this media to work on systems with those    *"
'Print #ff, "echo * mitigations applied. Without this update, this media may not boot on systems with *"
'Print #ff, "echo * those mitigations. If this disk will not boot on some systems, please do this:    *"
'Print #ff, "echo *                                                                                   *"
'Print #ff, "echo * 1) Rerun this program on a system with those mitigations installed and perform a  *"
'Print #ff, "echo *    REFRESH operation to refresh the boot information on this disk.                *"
'Print #ff, "echo *                                                                                   *"
'Print #ff, "echo * 2) If this program asks if you want to use a previously created BOOT.WIM file,    *"
'Print #ff, "echo *    choose NOT to use a previously created file.                                   *"
'Print #ff, "echo *************************************************************************************"
'Print #ff, "echo."
'Print #ff, "REM Check to make sure that the first of two mitigations are applied to this system."
'Print #ff, "REM The first mitigation adds the "; Chr$(34); "Windows UEFI CA 2023"; Chr$(34); " certificate to the UEFI "; Chr$(34); "Secure Boot Signature Database"; Chr$(34); " (DB)."
'Print #ff, "REM By adding this certificate to the DB, the device firmware will trust boot applications signed by this certificate."
'Print #ff, ""
'Print #ff, "for /f %%a in ('powershell "; Chr$(34); "[System.Text.Encoding]::ASCII.GetString((Get-SecureBootUEFI db).bytes) -match 'Windows UEFI CA 2023'"; Chr$(34); "') do set "; Chr$(34); "PowerShellOutput=%%a"; Chr$(34); ""
'Print #ff, ""
'Print #ff, "if "; Chr$(34); "%PowerShellOutput%"; Chr$(34); "=="; Chr$(34); "True"; Chr$(34); " ("
'Print #ff, "goto CheckCondition2"
'Print #ff, ") else ("
'Print #ff, "goto NotInstalled"
'Print #ff, ")"
'Print #ff, ""
'Print #ff, ":CheckCondition2"
'Print #ff, ""
'Print #ff, "REM Check to make sure that the second of two mitigations are applied to this system."
'Print #ff, "REM The UEFI Forbidden List (DBX) is used to block untrusted UEFI modules from loading. The second mitigation updates"
'Print #ff, "REM the DBX by adding the "; Chr$(34); "Windows Production CA 2011"; Chr$(34); " certificate to the DBX. This will cause all boot managers signed by"
'Print #ff, "REM this certificate to no longer be trusted."
'Print #ff, ""
'Print #ff, "for /f %%a in ('powershell "; Chr$(34); "[System.Text.Encoding]::ASCII.GetString((Get-SecureBootUEFI dbx).bytes) -match 'Microsoft Windows Production PCA 2011'"; Chr$(34); "') do set "; Chr$(34); "PowerShellOutput=%%a"; Chr$(34); ""
'Print #ff, ""
'Print #ff, "if "; Chr$(34); "%PowerShellOutput%"; Chr$(34); "=="; Chr$(34); "True"; Chr$(34); " ("
'Print #ff, "goto Condition2True"
'Print #ff, ") else ("
'Print #ff, "goto NotInstalled"
'Print #ff, ")"
'Print #ff, ""
'Print #ff, ":Condition2True"
'Print #ff, ""
'Print #ff, "REM We have verified that the mitigations for the BlackLotus UEFI Bootkit are installed on this system. We will now update"
'Print #ff, "REM the boot media to ensure that it can be successfully booted on this system."
'Print #ff, ""
'Print #ff, "REM Make sure that the files on the destination disk are not read only"
'Print #ff, ""
'Print #ff, "attrib -r "; Letter$(1); ":\*.* /s /d > NUL 2>&1"
'Print #ff, "COPY /Y /B "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD.BAK > NUL 2>&1"
'Print #ff, "bcdboot c:\windows /f UEFI /s "; Letter$(1); ": /bootex > NUL"
'Print #ff, "COPY /Y /B "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD.BAK "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD > NUL"
'Print #ff, "echo This system DOES have the mitigations installed. Media has been patched."
'Print #ff, "echo."
'Print #ff, "goto DonePatching"
'Print #ff, ""
'Print #ff, ":NotInstalled"
'Print #ff, ""
'Print #ff, "REM We arrive here if the mitigations are not installed on this system or when updates are doing being installed."
'Print #ff, "echo This system DOES NOT have the mitigations installed. Media has NOT been patched."
'Print #ff, ""
'Print #ff, ":DonePatching"
'Print #ff, ""
'Print #ff, "REM Done with routine to patch for Secure Boot mitigations"
'Print #ff, "pause"
Close #ff
Shell "TEMP.BAT"
If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

' Check for the existance of a file named "WIM_File_Copy_Error_1.txt" or "WIM_File_Copy_Error_2.txt. If such
' a file exists, it indicates that there was an error copying files with the above batch file. In that case,
' take the following actions:
'
' 1) Display a warning to the user.
' 2) Delete the "WIM_File_Copy_Error.txt" file.
' 3) Abort this routine and return to the start of the program.

If _FileExists("WIM_File_Copy_Error_1.txt") Then
    Shell _Hide "del WIM_File_Copy_Error_1.txt > NUL 2>&1"
    Cls
    Print
    Color 14, 4: Print "WARNING!";: Color 15: Print " There was an error copying files to partition 1. This usually indicates that there was not enough space on the"
    Print "destination. Please correct this situation and run this routine again."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

If _FileExists("WIM_File_Copy_Error_2.txt") Then
    Shell _Hide "del WIM_File_Copy_Error_2.txt > NUL 2>&1"
    Cls
    Print
    Color 14, 4: Print "WARNING!";: Color 15: Print " There was an error copying files to partition 2. This usually indicates that there was not enough space on the"
    Print "destination. Please correct this situation and run this routine again."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

' Making the file ei.cfg on the partition 2, in the sources folder. If an ei.cfg already exists, leave it alone

If CreateEiCfg$ = "Y" Then
    Temp$ = Letter$(2) + ":\sources\ei.cfg"

    If Not (_FileExists(Temp$)) Then
        ff = FreeFile
        Open (Temp$) For Output As #ff
        Print #ff, "[CHANNEL]"
        Print #ff, "Retail"
        Close #ff
    End If
End If

GoTo OpsCompleted

DualArchitecture:

ff = FreeFile
Open "TEMP.BAT" For Output As #ff
Print #ff, "@echo off"
Print #ff, "del WIM_File_Copy_Error_1.txt > NUL 2>&1"
Print #ff, "del WIM_File_Copy_Error_2.txt > NUL 2>&1"
Print #ff, "echo."
Print #ff, "echo *************************************************************"
Print #ff, "echo * Copying files. Be aware that this can take quite a while, *"
Print #ff, "echo * especially on the 2nd partition and with slower media.    *"
Print #ff, "echo * Please be patient and allow this process to finish.       *"
Print #ff, "echo *************************************************************"
Print #ff, "echo."
Print #ff, "echo Copying files to partition #1"
Print #ff, "set CurrentPartition=1"

If ExcludeAutounattend$ = "N" Then
    Print #ff, "robocopy "; CDROM$; "\ "; Letter$(1); ":\ /mir /xd sources /njs /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Else
    Print #ff, "robocopy "; CDROM$; "\ "; Letter$(1); ":\ /mir /xf autounattend.xml /xd sources /njs /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
End If

Print #ff, "robocopy "; CDROM$; "\x64\sources "; Letter$(1); ":\x64\sources boot.wim /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x86\sources "; Letter$(1); ":\x86\sources boot.wim /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "echo Copying files to partition #2"
Print #ff, "set CurrentPartition=2"
Print #ff, "robocopy "; CDROM$; "\x64\sources "; Letter$(2); ":\x64\sources /mir /njh /njs /xf boot.wim /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x86\sources "; Letter$(2); ":\x86\sources /mir /njh /njs /xf boot.wim /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x64\boot "; Letter$(2); ":\x64\boot /mir /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x86\boot "; Letter$(2); ":\x86\boot /mir /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x64\efi "; Letter$(2); ":\x64\efi /mir /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x86\efi "; Letter$(2); ":\x86\efi /mir /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x64\support "; Letter$(2); ":\x64\support /mir /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x86\support "; Letter$(2); ":\x86\support /mir /njh /njs /256 /r:0 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "goto cleanup"
Print #ff, ":HandleError"
Print #ff, "echo An error occurred > WIM_File_Copy_Error_%CurrentPartition%.txt"
Print #ff, ":cleanup"
Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'"; MakeBootableSourceISO$; "'"; Chr$(34); Chr$(34)

' Commenting out the section below that patches for Secure Boot. It is no longer needed with the ADK and Windows
' PE release of May, 2024

'Print #ff, "cls"
'Print #ff, "REM Start of routine to patch for Secure Boot mitigations"
'Print #ff, ""
'Print #ff, "echo *************************************************************************************"
'Print #ff, "echo *                              Applying Secure Boot Fixes                           *"
'Print #ff, "echo *                                                                                   *"
'Print #ff, "echo * Checking to see if this system has BlackLotus UEFI Bootkit mitigations installed. *"
'Print #ff, "echo * If it does, we will use this to patch this media to work on systems with those    *"
'Print #ff, "echo * mitigations applied. Without this update, this media may not boot on systems with *"
'Print #ff, "echo * those mitigations. If this disk will not boot on some systems, please do this:    *"
'Print #ff, "echo *                                                                                   *"
'Print #ff, "echo * 1) Rerun this program on a system with those mitigations installed and perform a  *"
'Print #ff, "echo *    REFRESH operation to refresh the boot information on this disk.                *"
'Print #ff, "echo *                                                                                   *"
'Print #ff, "echo * 2) If this program asks if you want to use a previously created BOOT.WIM file,    *"
'Print #ff, "echo *    choose NOT to use a previously created file.                                   *"
'Print #ff, "echo *************************************************************************************"
'Print #ff, "echo."
'Print #ff, "REM Check to make sure that the first of two mitigations are applied to this system."
'Print #ff, "REM The first mitigation adds the "; Chr$(34); "Windows UEFI CA 2023"; Chr$(34); " certificate to the UEFI "; Chr$(34); "Secure Boot Signature Database"; Chr$(34); " (DB)."
'Print #ff, "REM By adding this certificate to the DB, the device firmware will trust boot applications signed by this certificate."
'Print #ff, ""
'Print #ff, "for /f %%a in ('powershell "; Chr$(34); "[System.Text.Encoding]::ASCII.GetString((Get-SecureBootUEFI db).bytes) -match 'Windows UEFI CA 2023'"; Chr$(34); "') do set "; Chr$(34); "PowerShellOutput=%%a"; Chr$(34); ""
'Print #ff, ""
'Print #ff, "if "; Chr$(34); "%PowerShellOutput%"; Chr$(34); "=="; Chr$(34); "True"; Chr$(34); " ("
'Print #ff, "goto CheckCondition2"
'Print #ff, ") else ("
'Print #ff, "goto NotInstalled"
'Print #ff, ")"
'Print #ff, ""
'Print #ff, ":CheckCondition2"
'Print #ff, ""
'Print #ff, "REM Check to make sure that the second of two mitigations are applied to this system."
'Print #ff, "REM The UEFI Forbidden List (DBX) is used to block untrusted UEFI modules from loading. The second mitigation updates"
'Print #ff, "REM the DBX by adding the "; Chr$(34); "Windows Production CA 2011"; Chr$(34); " certificate to the DBX. This will cause all boot managers signed by"
'Print #ff, "REM this certificate to no longer be trusted."
'Print #ff, ""
'Print #ff, "for /f %%a in ('powershell "; Chr$(34); "[System.Text.Encoding]::ASCII.GetString((Get-SecureBootUEFI dbx).bytes) -match 'Microsoft Windows Production PCA 2011'"; Chr$(34); "') do set "; Chr$(34); "PowerShellOutput=%%a"; Chr$(34); ""
'Print #ff, ""
'Print #ff, "if "; Chr$(34); "%PowerShellOutput%"; Chr$(34); "=="; Chr$(34); "True"; Chr$(34); " ("
'Print #ff, "goto Condition2True"
'Print #ff, ") else ("
'Print #ff, "goto NotInstalled"
'Print #ff, ")"
'Print #ff, ""
'Print #ff, ":Condition2True"
'Print #ff, ""
'Print #ff, "REM We have verified that the mitigations for the BlackLotus UEFI Bootkit are installed on this system. We will now update"
'Print #ff, "REM the boot media to ensure that it can be successfully booted on this system."
'Print #ff, ""
'Print #ff, "REM Make sure that the files on the destination disk are not read only"
'Print #ff, ""
'Print #ff, "attrib -r "; Letter$(1); ":\*.* /s /d > NUL 2>&1"
'Print #ff, "COPY /Y /B "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD.BAK > NUL 2>&1"
'Print #ff, "bcdboot c:\windows /f UEFI /s "; Letter$(1); ": /bootex > NUL"
'Print #ff, "COPY /Y /B "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD.BAK "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD > NUL"
'Print #ff, "echo This system DOES have the mitigations installed. Media has been patched."
'Print #ff, "echo."
'Print #ff, "goto DonePatching"
'Print #ff, ""
'Print #ff, ":NotInstalled"
'Print #ff, ""
'Print #ff, "REM We arrive here if the mitigations are not installed on this system or when updates are doing being installed."
'Print #ff, "echo This system DOES NOT have the mitigations installed. Media has NOT been patched."
'Print #ff, ""
'Print #ff, ":DonePatching"
'Print #ff, ""
'Print #ff, "REM Done with routine to patch for Secure Boot mitigations"

Close #ff
Shell "TEMP.BAT"
If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

' Check for the existance of a file named "WIM_File_Copy_Error_1.txt" or "WIM_File_Copy_Error_2.txt". If such a file exists,
' it indicates that there was an error copying files with the above batch file. In that case, take the following actions:
'
' 1) Display a warning to the user.
' 2) Delete the "WIM_File_Copy_Error_x.txt" file.
' 3) Abort this routine and return to the start of the program.

If _FileExists("WIM_File_Copy_Error_1.txt") Then
    Shell _Hide "del WIM_File_Copy_Error_1.txt > NUL 2>&1"
    Cls
    Print
    Color 14, 4: Print "WARNING!";: Color 15: Print " There was an error copying files to partition 1. This usually indicates that there was not enough space on the"
    Print "destination. Please correct this situation and run this routine again."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

If _FileExists("WIM_File_Copy_Error_2.txt") Then
    Shell _Hide "del WIM_File_Copy_Error_2.txt > NUL 2>&1"
    Cls
    Print
    Color 14, 4: Print "WARNING!";: Color 15: Print " There was an error copying files to partition 2. This usually indicates that there was not enough space on the"
    Print "destination. Please correct this situation and run this routine again."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

' Making the file ei.cfg on the partition 2, in the \x64\sources and \x86\sources folders.

If CreateEiCfg$ = "Y" Then
    Temp$ = Letter$(2) + ":\x64\sources\ei.cfg"

    If Not (_FileExists(Temp$)) Then
        ff = FreeFile
        Open (Temp$) For Output As #ff
        Print #ff, "[CHANNEL]"
        Print #ff, "Retail"
        Close #ff
    End If

    Temp$ = Letter$(2) + ":\x86\sources\ei.cfg"

    If Not (_FileExists(Temp$)) Then
        ff = FreeFile
        Open (Temp$) For Output As #ff
        Print #ff, "[CHANNEL]"
        Print #ff, "Retail"
        Close #ff
    End If
End If

OpsCompleted:

' All operations are complete.

Cls

' In order to display a more descriptive closing message, we need to know whether the user chose to wipe the entire disk
' or whether they simply chose to refresh the the image on a disk. The variable "WipeOrRefresh" will be set to 1 if they
' elected to wipe the disk, or it will be 2 if they decided to perform a refresh.

Select Case WipeOrRefresh
    Case 1
        Print "The disk that you selected was prepared and 2 partitions were created to make your image(s) bootable."

        If AddPart$ = "Y" Then
            Print
            Print Str$(AdditionalPartitions); " additional partition(s) were also created."
        End If

    Case 2
        Print "The boot volume and supporting second volume have been refreshed. If the disk contains other volumes, these"
        Print "were not altered in any way."
End Select
Pause

ChDir ProgramStartDir$: GoTo BeginProgram

CreateMultiImageDisk:

If WinPEFound = 0 Then
    Cls
    Color 14, 4: Print "WARNING!": Color 15
    Print
    Print "We did not find the Windows PE add-on for the ADK installed on your system. This feature requires that component."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

' User needs to select a disk to use. After calling the SelectDisk routine, the disk ID will be stored in the DiskID variable.

GetPE_TempPath:

' Create a temporary folder in which we can build our WinPE image.

WinPE_Temp$ = Environ$("TEMP") + "\WIM_TOOLS_TEMP"

CleanPath WinPE_Temp$
WinPE_Temp$ = Temp$

' Test to see if the folder already exists. If not, create it as a test, then delete it.

If _DirExists(WinPE_Temp$) Then
    Cls
    Cmd$ = "RMDIR /s /q " + WinPE_Temp$
    Shell _Hide Cmd$
End If

Cmd$ = "MD " + Chr$(34) + WinPE_Temp$ + Chr$(34) + " > NUL"
Shell Cmd$

If _DirExists(WinPE_Temp$) Then
    Cmd$ = "rd /S /Q " + Chr$(34) + WinPE_Temp$ + Chr$(34) + " > NUL"
    Shell Cmd$
Else
    Cls
    Print "That directory is not valid. Please try again."
    Pause
    GoTo GetPE_TempPath
End If

' Creating the files we need for this project: Create_Disk.bat, startnet.cmd, and Config_UFD.bat
' Start with Create_Disk.bat

ff = FreeFile
Open "Create_Disk.bat" For Output As #ff

Print #ff, "@echo off"
Print #ff, "setlocal enabledelayedexpansion"
Print #ff, "setlocal enableextensions"
Print #ff, "cd /d %~dp0"
Print #ff, "cls"
Print #ff, ""
Print #ff, "set ADK_Path="; ADKLocation$; "Assessment and Deployment Kit"
Print #ff, "set PE_Temp="; WinPE_Temp$
Print #ff, "set Vol1Size=2500"
Print #ff, ""
Print #ff, "REM Do not change the following variables. They build upon the variable "; Chr$(34); "ADK_Path"; Chr$(34); ", which you CAN change."
Print #ff, ""
Print #ff, "set PE_Path=%ADK_Path%\Windows Preinstallation Environment"
Print #ff, "set DiskID="; LTrim$(Str$(DiskID))
Print #ff, ""
Print #ff, "REM Set the Deployment and Imaging Tools Environment"
Print #ff, ""
Print #ff, "pushd %ADK_Path%\Deployment Tools"
Print #ff, "call DandISetEnv.bat"
Print #ff, "popd"
Print #ff, ""
Print #ff, ":AskUseSavedWIM"
Print #ff, ""
Print #ff, "cls"
Print #ff, ":: Determine if the user has an already modified boot.wim file that they want to use"
Print #ff, ""
Print #ff, "echo We will now customize Windows PE for this project. If you have previously saved a copy of Windows PE (boot.wim)"
Print #ff, "echo that was created with this program you can use it to save the time needed to customize that file."
Print #ff, "echo."
Print #ff, "echo NOTE: If you choose not to use a previously created boot.wim file, you will asked if you wish to add any boot"
Print #ff, "echo critical drivers to the boot.wim. These are drivers such as storage drivers, touchpad drivers, etc. that are"
Print #ff, "echo necessary for proper operation while in Windows setup. Again, this is only an option if you do not use a"
Print #ff, "echo previously created boot.wim file."
Print #ff, "echo."
Print #ff, "set /p UsePreviousWIM="; Chr$(34); "Do you want to use a previously created boot.wim file? "; Chr$(34)
Print #ff, "if [%UsePreviousWIM%]==[] goto AskUseSavedWIM"
Print #ff, "set UsePreviousWIM=%UsePreviousWIM:~0,1%"
Print #ff, "call ::TOUPPERCASE UsePreviousWIM"
Print #ff, "if %UsePreviousWIM%==Y goto PromptForWIM"
Print #ff, "if %UsePreviousWIM% NEQ N goto AskUseSavedWIM"
Print #ff, "goto CreateNewWIM"
Print #ff, ""
Print #ff, ":PromptForWIM"
Print #ff, ""
Print #ff, "cls"
Print #ff, "echo Please place the boot.wim file to use on your desktop now and then press any key to continue."
Print #ff, "echo."
Print #ff, "pause"
Print #ff, ""
Print #ff, ":PromptForWIMagain"
Print #ff, ""
Print #ff, "cls"
Print #ff, ""
Print #ff, ":: If the variable UsePreviousWIM is set to Y, then the CreateNewWIM will use that WIM rather than create a new one"
Print #ff, ""
Print #ff, "if exist "; Chr$(34); "C:\Users\%username%\Desktop\boot.wim"; Chr$(34); " goto CreateNewWIM"
Print #ff, ""
Print #ff, ":: The following actions are taken when a boot.wim is not found on the desktop."
Print #ff, ""
Print #ff, "echo A boot.wim file was not found on your desktop. You now have two options:"
Print #ff, "echo."
Print #ff, "echo 1) Place the boot.wim file to be used on the desktop and press ENTER."
Print #ff, "echo 2) Type the word NEW and press ENTER. This will make the program create a new boot.wim file."
Print #ff, "echo."
Print #ff, "set /p NoWIMfoundAction="; Chr$(34); "Press ENTER, or type NEW and then press ENTER. "; Chr$(34)
Print #ff, "if [%NoWIMfoundAction%]==[] goto PromptForWIMagain"
Print #ff, "call ::TOUPPERCASE NoWIMfoundAction"
Print #ff, "if %NoWIMfoundAction%==NEW ("
Print #ff, "set UsePreviousWIM=N"
Print #ff, "goto CreateNewWIM"
Print #ff, ")"
Print #ff, "goto PromptForWIMagain"
Print #ff, ""
Print #ff, ":CreateNewWIM"
Print #ff, ""
Print #ff, "cls"
Print #ff, "echo ::::::::::::::::::::::::::::::"
Print #ff, "echo :: Copying Windows PE Files ::"
Print #ff, "echo ::::::::::::::::::::::::::::::"
Print #ff, "echo."
Print #ff, ""
Print #ff, "cmd /c "; Chr$(34); Chr$(34); "%PE_Path%\copype"; Chr$(34); " amd64 "; Chr$(34); "%PE_Temp%"; Chr$(34); Chr$(34); " > NUL"
Print #ff, ""
Print #ff, "if %UsePreviousWIM%==Y ("
Print #ff, "copy /b "; Chr$(34); "C:\Users\%username%\Desktop\boot.wim"; Chr$(34); " "; Chr$(34); "%PE_Temp%\media\sources"; Chr$(34); " /Y > NUL"
Print #ff, "goto CreateDisk"
Print #ff, ")"
Print #ff, ""
Print #ff, "echo ::::::::::::::::::::::::::::::::::::::::"
Print #ff, "echo :: Mounting the Windows PE Image File ::"
Print #ff, "echo ::::::::::::::::::::::::::::::::::::::::"
Print #ff, "echo."
Print #ff, ""
Print #ff, "DISM /Mount-Image /ImageFile:"; Chr$(34); "%PE_Temp%\media\sources\boot.wim"; Chr$(34); " /Index:1 /MountDir:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " > NUL"
Print #ff, ""
Print #ff, "REM Add packages to Windows PE"
Print #ff, ""
Print #ff, "CLS"
Print #ff, "echo Adding package 1 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\WinPE-WMI.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 2 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\en-us\WinPE-WMI_en-us.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 3 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\WinPE-NetFX.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 4 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\en-us\WinPE-NetFX_en-us.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 5 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\WinPE-Scripting.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 6 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\en-us\WinPE-Scripting_en-us.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 7 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\WinPE-PowerShell.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 8 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\en-us\WinPE-PowerShell_en-us.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 9 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\WinPE-StorageWMI.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 10 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\en-us\WinPE-StorageWMI_en-us.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 11 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\WinPE-DismCmdlets.cab"; Chr$(34); " > NUL"
Print #ff, "CLS"
Print #ff, "echo Adding package 12 of 12 to Windows PE"
Print #ff, "Dism /Add-Package /Image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /PackagePath:"; Chr$(34); "%PE_Path%\amd64\WinPE_OCs\en-us\WinPE-DismCmdlets_en-us.cab"; Chr$(34); " > NUL"
Print #ff, ""
Print #ff, "CLS"
Print #ff, "echo If you need to add any drivers to Windows PE we can install them now. As an example, suppose that you have a touchpad"
Print #ff, "echo that will not work properly without drivers or a storage device that is not accessible unless you load drivers. We"
Print #ff, "echo will automatically add any drivers that you place on your desktop. The drivers must be placed into a folder called"
Print #ff, "echo PE_Drivers in a structure like this:"
Print #ff, "echo."
Print #ff, "echo PE_Drivers"
Print #ff, "echo    x64"
Print #ff, "echo       Driver 1"
Print #ff, "echo       Driver 2"
Print #ff, "echo."
Print #ff, "echo The folders named Driver 1 and Driver 2 above can have any name you want and you can have as many as you"
Print #ff, "echo wish. The drivers need to be extracted (not in a .exe, .cab, .zip, etc.) and must have a .INF file. "
Print #ff, "echo."
Print #ff, "echo If you want to do this, prepare that folder now, and then press any key to continue. Otherwise, just hit"
Print #ff, "echo ENTER now to proceed."
Print #ff, "echo."
Print #ff, "pause"
Print #ff, "CLS"
Print #ff, "echo Adding drivers (if available). This may take a while."
Print #ff, "echo."
Print #ff, "DISM /image:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /Add-Driver /Driver:"; Chr$(34); "C:\Users\%username%\Desktop\PE_Drivers\x64"; Chr$(34); " /recurse >NUL"
Print #ff, "CLS"
Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, "echo :: Adding a Custom startnet.cmd File to Windows PE ::"
Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, "echo."
Print #ff, ""
Print #ff, "copy /B startnet.cmd "; Chr$(34); "%PE_Temp%\mount\windows\system32"; Chr$(34); " > NUL"
Print #ff, ""
Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, "echo :: Saving and Dismounting the Windows PE Image ::"
Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ""
Print #ff, "DISM /Unmount-Image /MountDir:"; Chr$(34); "%PE_Temp%\mount"; Chr$(34); " /Commit > NUL"
Print #ff, "echo."
Print #ff, ""
Print #ff, ":CreateDisk"
Print #ff, ""
Print #ff, "echo ::::::::::::::::::::::::::::::::::"
Print #ff, "echo :: Creating the Final Boot Disk ::"
Print #ff, "echo ::::::::::::::::::::::::::::::::::"
Print #ff, "echo."
Print #ff, ""
Print #ff, "if exist "; Chr$(34); "%temp%\WinPE_ISO_Image.ISO"; Chr$(34); " ("
Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'%temp%\WinPE_ISO_Image.ISO'"; Chr$(34); Chr$(34) + " > NUL"
Print #ff, "del "; Chr$(34); "%temp%\WinPE_ISO_Image.ISO"; Chr$(34)
Print #ff, ")"
Print #ff, ""
Print #ff, "cmd /c "; Chr$(34); Chr$(34); "%PE_Path%\MakeWinPEMedia"; Chr$(34); " /ISO "; Chr$(34); "%PE_Temp%"; Chr$(34); " %temp%\WinPE_ISO_Image.ISO"; Chr$(34); " > NUL 2>&1"
Print #ff, ""
Print #ff, "REM Mount the image."
Print #ff, ""
Print #ff, "powershell.exe -command "; Chr$(34); "Mount-DiskImage "; Chr$(34); "'%temp%\WinPE_ISO_Image.ISO'"; Chr$(34); " -PassThru | Get-Volume"; Chr$(34); " > MountInfo.txt"
Print #ff, ""
Print #ff, "REM Get drive letter (includes the colon)."
Print #ff, "REM We need to skip 3 lines in order to read the 4th line of text in the MountInfo.txt file."
Print #ff, ""
Print #ff, "for /F "; Chr$(34); "skip=3 delims="; Chr$(34); " %%a in (MountInfo.txt) do ("
Print #ff, "   set LineContents=%%a"
Print #ff, "   goto Evaluate"
Print #ff, ")"
Print #ff, ""
Print #ff, ":Evaluate"
Print #ff, ""
Print #ff, "REM We are done with MountInfo.txt. Delete it."
Print #ff, ""
Print #ff, "del MountInfo.txt"
Print #ff, "set LineContents=%LineContents:~0,1%"
Print #ff, "set MountedImageDriveLetter=%LineContents%:"
Print #ff, "robocopy %MountedImageDriveLetter%\ "; Letter$(1); ":\ /mir /R:0 /xf VOL1_M_MEDIA.WIM /xd "; Chr$(34); "System Volume Information"; Chr$(34); " $Recycle.bin > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError1"
Print #ff, "goto ProcessPar2"
Print #ff, ":HandleError1"
Print #ff, "echo An error occurred when copying files to partition 1. This is most likely due to not assigning sufficient space to "
Print #ff, "echo that partition. The media created by this program should not be considered reliable as a result of this issue. When"
Print #ff, "echo the program is done running, please run this routine again but try assigning some more space to partition 1."
Print #ff, "echo."
Print #ff, "pause"
Print #ff, "echo."
Print #ff, "goto Cleanup"
Print #ff, ":ProcessPar2"
Print #ff, "move /Y config_ufd.bat "; Letter$(2); ":\ > NUL"
Print #ff, "MD "; Chr$(34); Letter$(2); ":\ISO Images"; Chr$(34); " > NUL 2>&1"
Print #ff, "MD "; Chr$(34); Letter$(2); ":\Other"; Chr$(34); " > NUL 2>&1"
Print #ff, "MD "; Chr$(34); Letter$(2); ":\Answer Files"; Chr$(34); " > NUL 2>&1"
Print #ff, "echo. > "; Letter$(1); ":\VOL1_M_MEDIA.WIM"
Print #ff, "echo. > "; Letter$(2); ":\VOL2_M_MEDIA.WIM"
Print #ff, ""
Print #ff, ":Cleanup"
Print #ff, ""
Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'%temp%\WinPE_ISO_Image.ISO'"; Chr$(34); Chr$(34) + " > NUL"
Print #ff, "echo."
Print #ff, ""
Print #ff, ":AskSaveWIM"
Print #ff, ""
Print #ff, "cls"
Print #ff, "if %UsePreviousWIM%==Y goto ContinueCleanup"
Print #ff, "echo You have the option to save the boot.wim file that we created. This will allow you to use that file in the future"
Print #ff, "echo and save the time that it would normally take to build that file."
Print #ff, "echo."
Print #ff, "set /p SaveWim="; Chr$(34); "Do you want to save a copy of the boot.wim that we created for future use? "; Chr$(34)
Print #ff, ""
Print #ff, "if [%SaveWim%]==[] goto AskSaveWIM"
Print #ff, "set SaveWIM=%SaveWIM:~0,1%"
Print #ff, "call ::TOUPPERCASE SaveWIM"
Print #ff, "if %SaveWIM%==Y goto SaveWIMtoDesktop"
Print #ff, "if %SaveWim% NEQ N ("
Print #ff, "goto AskSaveWIM"
Print #ff, ") ELSE ("
Print #ff, "goto ContinueCleanup"
Print #ff, ")"
Print #ff, ""
Print #ff, ""
Print #ff, ":SaveWIMtoDesktop"
Print #ff, ""
Print #ff, "cls"
Print #ff, "echo We will save the boot.wim file to your desktop. If you already have a boot.wim file there and you want to keep it,"
Print #ff, "echo please move or rename the file now."
Print #ff, "echo."
Print #ff, "pause"
Print #ff, "cls"
Print #ff, "echo Copying the the boot.wim file to your desktop."
Print #ff, "copy /b "; Chr$(34); "%PE_Temp%\media\sources\boot.wim"; Chr$(34); " "; Chr$(34); "C:\Users\%username%\Desktop"; Chr$(34); " > NUL"
Print #ff, ""
Print #ff, ":ContinueCleanup"
Print #ff, ""
Print #ff, "RD /S /Q "; Chr$(34); "%PE_Temp%"; Chr$(34); " > NUL"
Print #ff, "del "; Chr$(34); "%temp%\WinPE_ISO_Image.ISO"; Chr$(34); " > NUL"

' Commenting out the section below that patches for Secure Boot. It is no longer needed with the ADK and Windows
' PE release of May, 2024

'Print #ff, "cls"
'Print #ff, "REM Start of routine to patch for Secure Boot mitigations"
'Print #ff, ""
'Print #ff, "echo *************************************************************************************"
'Print #ff, "echo *                              Applying Secure Boot Fixes                           *"
'Print #ff, "echo *                                                                                   *"
'Print #ff, "echo * Checking to see if this system has BlackLotus UEFI Bootkit mitigations installed. *"
'Print #ff, "echo * If it does, we will use this to patch this media to work on systems with those    *"
'Print #ff, "echo * mitigations applied. Without this update, this media may not boot on systems with *"
'Print #ff, "echo * those mitigations. If this disk will not boot on some systems, please do this:    *"
'Print #ff, "echo *                                                                                   *"
'Print #ff, "echo * 1) Rerun this program on a system with those mitigations installed and perform a  *"
'Print #ff, "echo *    REFRESH operation to refresh the boot information on this disk.                *"
'Print #ff, "echo *                                                                                   *"
'Print #ff, "echo * 2) If this program asks if you want to use a previously created BOOT.WIM file,    *"
'Print #ff, "echo *    choose NOT to use a previously created file.                                   *"
'Print #ff, "echo *************************************************************************************"
'Print #ff, "echo."
'Print #ff, "REM Check to make sure that the first of two mitigations are applied to this system."
'Print #ff, "REM The first mitigation adds the "; Chr$(34); "Windows UEFI CA 2023"; Chr$(34); " certificate to the UEFI "; Chr$(34); "Secure Boot Signature Database"; Chr$(34); " (DB)."
'Print #ff, "REM By adding this certificate to the DB, the device firmware will trust boot applications signed by this certificate."
'Print #ff, ""
'Print #ff, "for /f %%a in ('powershell "; Chr$(34); "[System.Text.Encoding]::ASCII.GetString((Get-SecureBootUEFI db).bytes) -match 'Windows UEFI CA 2023'"; Chr$(34); "') do set "; Chr$(34); "PowerShellOutput=%%a"; Chr$(34); ""
'Print #ff, ""
'Print #ff, "if "; Chr$(34); "%PowerShellOutput%"; Chr$(34); "=="; Chr$(34); "True"; Chr$(34); " ("
'Print #ff, "goto CheckCondition2"
'Print #ff, ") else ("
'Print #ff, "goto NotInstalled"
'Print #ff, ")"
'Print #ff, ""
'Print #ff, ":CheckCondition2"
'Print #ff, ""
'Print #ff, "REM Check to make sure that the second of two mitigations are applied to this system."
'Print #ff, "REM The UEFI Forbidden List (DBX) is used to block untrusted UEFI modules from loading. The second mitigation updates"
'Print #ff, "REM the DBX by adding the "; Chr$(34); "Windows Production CA 2011"; Chr$(34); " certificate to the DBX. This will cause all boot managers signed by"
'Print #ff, "REM this certificate to no longer be trusted."
'Print #ff, ""
'Print #ff, "for /f %%a in ('powershell "; Chr$(34); "[System.Text.Encoding]::ASCII.GetString((Get-SecureBootUEFI dbx).bytes) -match 'Microsoft Windows Production PCA 2011'"; Chr$(34); "') do set "; Chr$(34); "PowerShellOutput=%%a"; Chr$(34); ""
'Print #ff, ""
'Print #ff, "if "; Chr$(34); "%PowerShellOutput%"; Chr$(34); "=="; Chr$(34); "True"; Chr$(34); " ("
'Print #ff, "goto Condition2True"
'Print #ff, ") else ("
'Print #ff, "goto NotInstalled"
'Print #ff, ")"
'Print #ff, ""
'Print #ff, ":Condition2True"
'Print #ff, ""
'Print #ff, "REM We have verified that the mitigations for the BlackLotus UEFI Bootkit are installed on this system. We will now update"
'Print #ff, "REM the boot media to ensure that it can be successfully booted on this system."
'Print #ff, ""
'Print #ff, "REM Make sure that the files on the destination disk are not read only"
'Print #ff, ""
'Print #ff, "attrib -r "; Letter$(1); ":\*.* /s /d > NUL 2>&1"
'Print #ff, "COPY /Y /B "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD.BAK > NUL 2>&1"
'Print #ff, "bcdboot c:\windows /f UEFI /s "; Letter$(1); ": /bootex > NUL"
'Print #ff, "COPY /Y /B "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD.BAK "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD > NUL"
'Print #ff, "echo This system DOES have the mitigations installed. Media has been patched."
'Print #ff, "echo."
'Print #ff, "goto DonePatching"
'Print #ff, ""
'Print #ff, ":NotInstalled"
'Print #ff, ""
'Print #ff, "REM We arrive here if the mitigations are not installed on this system or when updates are doing being installed."
'Print #ff, "echo This system DOES NOT have the mitigations installed. Media has NOT been patched."
'Print #ff, ""
'Print #ff, ":DonePatching"
'Print #ff, ""
'Print #ff, "REM Done with routine to patch for Secure Boot mitigations"
'Print #ff, "pause"

Print #ff, "cls"
Print #ff, "goto finish"
Print #ff, ""
Print #ff, ""
Print #ff, ":::::::::::::::"
Print #ff, ":: FUNCTIONS ::"
Print #ff, ":::::::::::::::"
Print #ff, ""
Print #ff, ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ":: The following is a function that can be called to convert ::"
Print #ff, ":: the contents of a variable to uppercase characters.       ::"
Print #ff, "::                                                           ::"
Print #ff, ":: To use this function, call this function and pass it the  ::"
Print #ff, ":: name of the variable to convert. In the example below we  ::"
Print #ff, ":: are passing str as the variable name.                     ::"
Print #ff, "::                                                           ::"
Print #ff, ":: Example: call ::TOUPPERCASE str                           ::"
Print #ff, ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ""
Print #ff, ":TOUPPERCASE"
Print #ff, ""
Print #ff, "if not defined %~1 exit /b"
Print #ff, "for %%a in ("; Chr$(34); "a=A"; Chr$(34); " "; Chr$(34); "b=B"; Chr$(34); " "; Chr$(34); "c=C"; Chr$(34); " "; Chr$(34); "d=D"; Chr$(34); " "; Chr$(34); "e=E"; Chr$(34); " "; Chr$(34); "f=F"; Chr$(34); " "; Chr$(34); "g=G"; Chr$(34); " "; Chr$(34); "h=H"; Chr$(34); " "; Chr$(34); "i=I"; Chr$(34); " "; Chr$(34); "j=J"; Chr$(34); " "; Chr$(34); "k=K"; Chr$(34); " "; Chr$(34); "l=L"; Chr$(34); " "; Chr$(34); "m=M"; Chr$(34); " "; Chr$(34); "n=N"; Chr$(34); " "; Chr$(34); "o=O"; Chr$(34); " "; Chr$(34); "p=P"; Chr$(34); " "; Chr$(34); "q=Q"; Chr$(34); " "; Chr$(34); "r=R"; Chr$(34); " "; Chr$(34); "s=S"; Chr$(34); " "; Chr$(34); "t=T"; Chr$(34); " "; Chr$(34); "u=U"; Chr$(34); " "; Chr$(34); "v=V"; Chr$(34); " "; Chr$(34); "w=W"; Chr$(34); " "; Chr$(34); "x=X"; Chr$(34); " "; Chr$(34); "y=Y"; Chr$(34); " "; Chr$(34); "z=Z"; Chr$(34); " "; Chr$(34); "ä=Ä"; Chr$(34); " "; Chr$(34); "ö=Ö"; Chr$(34); " "; Chr$(34); "ü=Ü"; Chr$(34); ") do ("
Print #ff, "call set %~1=%%%~1:%%~a%%"
Print #ff, ")"
Print #ff, "goto :eof"
Print #ff, ""
Print #ff, ""
Print #ff, "::::::::::::::::::::::"
Print #ff, ":: END OF FUNCTIONS ::"
Print #ff, "::::::::::::::::::::::"
Print #ff, ""
Print #ff, ""
Print #ff, ":Finish"
Print #ff, ""
Close #ff

' Create the startnet.cmd file and Config_UFD.bat

For x = 1 To 2

    Select Case x
        Case 1
            ff = FreeFile
            Open "startnet.cmd" For Output As #ff
        Case 2
            ff = FreeFile
            Open "Config_UFD.bat" For Output As #ff
    End Select

    Print #ff, "@echo off"
    Print #ff, "setlocal enabledelayedexpansion"
    Print #ff, "setlocal enableextensions"
    Print #ff, "cd /d %~dp0"
    Print #ff, ""
    Print #ff, "REM We need to set an initial value for the variable af_FileNameOnly to prevent problems later."
    Print #ff, ""
    Print #ff, "set af_FileNameOnly="
    Print #ff,

    ' The following lines are only needed in the startnet.cmd file, not in config_ufd.bat

    If x = 1 Then
        Print #ff, "REM Initialize networking and enable the High Performance power plan"
        Print #ff, ""
        Print #ff, "wpeinit"
        Print #ff, "powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
        Print #ff, ""
    End If

    ' The following lines are only needed in the config_ufd.bat file

    If x = 2 Then
        Print #ff, ""
        Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #ff, ":: Check to see if this batch file is being run as Administrator. If it is not, then rerun the batch file ::"
        Print #ff, ":: automatically as admin and terminate the initial instance of the batch file.                           ::"
        Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #ff, ""
        Print #ff, ""
        Print #ff, "(Fsutil Dirty Query %SystemDrive%>Nul)||(PowerShell start "; Chr$(34); ""; Chr$(34); ""; Chr$(34); "%~f0"; Chr$(34); ""; Chr$(34); ""; Chr$(34); " -verb RunAs & Exit /B) > NUL 2>&1"
        Print #ff, ""
        Print #ff, ""
        Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #ff, ":: End Routine to check if being run as Admin ::"
        Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #ff, ""
        Print #ff, ""
        Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #ff, ":: Change the console mode to 120 columns wide by 25 lines high ::"
        Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #ff, ""
        Print #ff, ""
        Print #ff, "mode con: cols=120 lines=25"
        Print #ff, ""
    End If

    Print #ff, "REM Need to get the drive letters for the 2 volumes of the boot media"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Retrieving drive letters. Please standby..."
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "REM Finding drive letter for Volume 1"
    Print #ff, ""
    Print #ff, "FOR %%a IN (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do ("
    Print #ff, "IF exist %%a:\VOL1_M_MEDIA.WIM ("
    Print #ff, "set Vol1=%%a"
    Print #ff, "goto Vol1Found"
    Print #ff, "   )"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "echo We could not find Volume 1. The program will now end."
    Print #ff, "pause"
    Print #ff, "goto End"
    Print #ff, ""
    Print #ff, ":Vol1Found"
    Print #ff, ""
    Print #ff, "REM Finding drive letter for Volume 2"
    Print #ff, ""
    Print #ff, "FOR %%a IN (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do ("
    Print #ff, "IF exist %%a:\VOL2_M_MEDIA.WIM ("
    Print #ff, "set Vol2=%%a"
    Print #ff, "goto Vol2Found"
    Print #ff, "   )"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "echo We could not find Volume 2. The program will now end."
    Print #ff, "pause"
    Print #ff, "goto End"
    Print #ff, ""
    Print #ff, ":Vol2Found"
    Print #ff, ""
    Print #ff, "if not exist %Vol2%:\Boot_Image.txt goto VerifyImages"
    Print #ff, ""
    Print #ff, ":AskRevert"
    Print #ff, ""
    Print #ff, "REM If we reach this point then the disk was previously configured. We will display what image"
    Print #ff, "REM it is configured to boot. If an answer file was selected we will also show that."
    Print #ff, ""
    Print #ff, "CLS"
    Print #ff, "echo **********************************"
    Print #ff, "echo * Your disk will boot this image *"
    Print #ff, "echo **********************************"
    Print #ff, "echo."
    Print #ff, "type %Vol2%:\Boot_Image.txt"
    Print #ff, "echo."
    Print #ff, "if NOT exist %Vol1%:\autounattend.xml goto NoAnsFileOnVol1"
    Print #ff, "echo *********************************"
    Print #ff, "echo * This answer file will be used *"
    Print #ff, "echo *********************************"
    Print #ff, "echo."
    Print #ff, "if not exist %Vol1%:\answer_file.txt ("
    Print #ff, "echo An answer file generated by the user or one manually copied by the user is present."
    Print #ff, "echo."
    Print #ff, ") ELSE ("
    Print #ff, "type %Vol1%:\Answer_File.txt"
    Print #ff, "echo."
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "REM The program places an answer file on Vol1, not Vol2. in theory, we should"
    Print #ff, "REM never encounter an answer file on Vol2, but for extra safety, we will"
    Print #ff, "REM report if we find one there. If we already found an answer file on Vol1,"
    Print #ff, "REM then there is no need to report it again so we skip checking for one on Vol2."
    Print #ff, ""
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: CAUTION^^! Your disk includes an unattended answer file ^(autounattend.xml^). If your ::"
    Print #ff, "echo :: system is configured to boot from the disk, Windows installation will begin       ::"
    Print #ff, "echo :: automatically. If your answer file is configured to wipe a disk^(s^), then this     ::"
    Print #ff, "echo :: will happen automatically with no warning.                                        ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, "goto Revert_Continue"
    Print #ff, ":NoAnsFileOnVol1"
    Print #ff, ""
    Print #ff, "if exist %Vol2%:\autounattend.xml ("
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: CAUTION^^! Your disk includes an unattended answer file ^(autounattend.xml^). If your ::"
    Print #ff, "echo :: system is configured to boot from the disk, Windows installation will begin       ::"
    Print #ff, "echo :: automatically. If your answer file is configured to wipe a disk^(s^), then this     ::"
    Print #ff, "echo :: will happen automatically with no warning.                                        ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, "echo Because the answer file used was placed onto the 2nd volume manually ^(not by using this program^), we are unable"
    Print #ff, "echo to provide any details about this answer file."
    Print #ff, "echo."
    Print #ff, "goto Revert_Continue"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, ":Revert_Continue"
    Print #ff, ""
    Print #ff, "set /P Revert="; Chr$(34); "Do you want to revert the disk back to the original state so that a new image can be selected? "; Chr$(34)
    Print #ff, "if [%Revert%]==[] goto AskRevert"
    Print #ff, "set Revert=%Revert:~0,1%"
    Print #ff, "call ::TOUPPERCASE Revert"
    Print #ff, "if %Revert%==Y goto Restore"
    Print #ff, "if %Revert% NEQ N goto AskRevert"
    Print #ff, "goto End"
    Print #ff, ""
    Print #ff, ":VerifyImages"
    Print #ff, ""
    Print #ff, "REM Verify that the ISO Images folder is not empty"
    Print #ff, ""
    Print #ff, "set cnt=0"
    Print #ff, "pushd "; Chr$(34); "%Vol2%:\ISO Images"; Chr$(34); ""
    Print #ff, "for %%a in (*.ISO) do set /a cnt+=1"
    Print #ff, "if %cnt%==0 ("
    Print #ff, "popd"
    Print #ff, "cls"
    Print #ff, "echo No ISO image files were found in "; Chr$(34); "%Vol2%:\ISO Images"; Chr$(34); ". Did you forget to place your images in that folder?"
    Print #ff, "echo."
    Print #ff, "echo Please correct this and then run this batch file again. The program will now end when you press any key."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto End"
    Print #ff, ")"
    Print #ff, "popd"
    Print #ff, ""
    Print #ff, "REM Ask user for the image to deploy"
    Print #ff, ""
    Print #ff, ":GetImageName"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, ""
    Print #ff, "echo Below is a list of available ISO image files:"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "echo [0] or [ENTER] Exit without making any image bootable"
    Print #ff, ""
    Print #ff, "set filechoice=0"
    Print #ff, "set count=0"
    Print #ff, "set "; Chr$(34); "choice_options="; Chr$(34); ""
    Print #ff, ""
    Print #ff, "for /F "; Chr$(34); "delims="; Chr$(34); " %%A in ('dir /a:-d /b "; Chr$(34); "%Vol2%:\ISO Images\*.iso"; Chr$(34); "') do ("
    Print #ff, ""
    Print #ff, "REM Increment the image file count"
    Print #ff, ""
    Print #ff, "set /a count+=1"
    Print #ff, ""
    Print #ff, "REM Add the file name to the options array"
    Print #ff, ""
    Print #ff, "set "; Chr$(34); "options[!count!]=%%A"; Chr$(34); ""
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "REM Add the image file name to an array"
    Print #ff, ""
    Print #ff, "for /L %%A in (1,1,!count!) do echo [%%A] !options[%%A]!"
    Print #ff, ""
    Print #ff, "REM Ask the user to select an image"
    Print #ff, ""
    Print #ff, "echo."
    Print #ff, "set /p filechoice="; Chr$(34); "Enter the number of the ISO image you wish to use: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, ""
    Print #ff, "if %filechoice% EQU 0 ("
    Print #ff, "goto End"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if %filechoice% LSS 1 ("
    Print #ff, "echo."
    Print #ff, "echo Please provide a valid response."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto GetImageName"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if %filechoice% GTR %count% ("
    Print #ff, "echo."
    Print #ff, "echo Please provide a valid response"
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto GetImageName"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "REM Set FileNameOnly to hold the name of the file without a path"
    Print #ff, "REM Set SourceISOImage to hold the full path, including the filename"
    Print #ff, ""
    Print #ff, "set FileNameOnly=!options[%filechoice%]!"
    Print #ff, "set SourceISOImage=%Vol2%:\ISO Images\!options[%filechoice%]!"
    Print #ff, ""
    Print #ff, "if not exist "; Chr$(34); "%SourceISOImage%"; Chr$(34); " ("
    Print #ff, "cls"
    Print #ff, "echo We could not find an image by that name. Please check the name and try again."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto GetImageName"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "REM Save the name of the selected image to a file called Boot_Image.txt"
    Print #ff, "echo %FileNameOnly% > %Vol2%:\Boot_Image.txt"
    Print #ff, ":VerifyAnswerFiles"
    Print #ff, ""
    Print #ff, "REM Verify that the Answer Files folder is not empty"
    Print #ff, ""
    Print #ff, "set cnt=0"
    Print #ff, "pushd "; Chr$(34); "%Vol2%:\Answer Files"; Chr$(34); ""
    Print #ff, "for %%a in (*.XML) do set /a cnt+=1"
    Print #ff, "if %cnt%==0 ("
    Print #ff, "popd"
    Print #ff, "goto DoneWithAnswerFiles"
    Print #ff, ")"
    Print #ff, "popd"
    Print #ff, ""
    Print #ff, "REM Ask user for the answer file to deploy"
    Print #ff, ""
    Print #ff, ":GetAnswerFile"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, ""
    Print #ff, "echo Below is a list of available answer files. Note that answer files are only effective with Windows images:"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "echo [0] or [ENTER] Exit without selecting an answer file or to generate a new answer file"
    Print #ff, ""
    Print #ff, "set af_filechoice=0"
    Print #ff, "set af_count=0"
    Print #ff, "set "; Chr$(34); "af_choice_options="; Chr$(34); ""
    Print #ff, ""
    Print #ff, "for /F "; Chr$(34); "delims="; Chr$(34); " %%A in ('dir /a:-d /b "; Chr$(34); "%Vol2%:\Answer Files\*.xml"; Chr$(34); "') do ("
    Print #ff, ""
    Print #ff, "REM Increment the answer file count"
    Print #ff, ""
    Print #ff, "set /a af_count+=1"
    Print #ff, ""
    Print #ff, "REM Add the file name to the options array"
    Print #ff, ""
    Print #ff, "set "; Chr$(34); "af_options[!af_count!]=%%A"; Chr$(34); ""
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "REM Add the answer file name to an array"
    Print #ff, ""
    Print #ff, "for /L %%A in (1,1,!af_count!) do echo [%%A] !af_options[%%A]!"
    Print #ff, ""
    Print #ff, "REM Ask the user to select an answer file"
    Print #ff, ""
    Print #ff, "echo."
    Print #ff, "set /p af_filechoice="; Chr$(34); "Enter the number of the answer file you wish to use: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, "if %af_filechoice% EQU 0 ("
    Print #ff, "goto DoneWithAnswerFiles"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if %af_filechoice% LSS 1 ("
    Print #ff, "echo."
    Print #ff, "echo Please provide a valid response."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto GetImageName"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if %af_filechoice% GTR %af_count% ("
    Print #ff, "echo."
    Print #ff, "echo Please provide a valid response"
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto GetAnswerFile"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "REM Set af_FileNameOnly to hold the name of the file without a path"
    Print #ff, "REM Set af_UserAnswerFile to hold the full path, including the filename"
    Print #ff, ""
    Print #ff, "set af_FileNameOnly=!af_options[%af_filechoice%]!"
    Print #ff, "set af_UserAnswerFile=%Vol2%:\Answer Files\!af_options[%af_filechoice%]!"
    Print #ff, ""
    Print #ff, "if not exist "; Chr$(34); "%af_UserAnswerFile%"; Chr$(34); " ("
    Print #ff, "cls"
    Print #ff, "echo We could not find an answer file by that name. Please check the name and try again."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto GetAnswerFile"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, ":DoneWithAnswerFiles"
    Print #ff, ""
    Print #ff, "REM If the user did not select a predefined answer file, then offer the choice to create an"
    Print #ff, "REM answer file on the fly."
    Print #ff, ""
    Print #ff, "if "; Chr$(34); "%af_filechoice%"; Chr$(34); "=="; Chr$(34); Chr$(34); " goto AskGen"
    Print #ff, "if NOT "; Chr$(34); "%af_filechoice%"; Chr$(34); "=="; Chr$(34); "0"; Chr$(34); " goto AnsGenDone"
    Print #ff, ""
    Print #ff, "REM The user did NOT select a predefined answer file. Offer the option to generate an answer file on the fly."
    Print #ff, ""
    Print #ff, ":AskGen"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "choice /C YN /N /M "; Chr$(34); "Do you want to create a new answer file now? Press Y for YES, N for NO:"; Chr$(34)
    Print #ff, "if errorlevel 2 goto AnsGenDone"
    Print #ff, "if errorlevel 1 goto GenerateAnsFile"
    Print #ff, ""
    Print #ff, ":GenerateAnsFile"
    Print #ff, "pushd %~dp0"
    Print #ff, ":RunInteractiveMode"
    Print #ff, ""
    Print #ff, ":: In interactive mode, we are not using the predefinbed variables so we will set initial variables to empty with"
    Print #ff, ":: a few exceptions so that the user can accept the default values."
    Print #ff, ""
    Print #ff, "set SystemType="
    Print #ff, "set EfiParSize=260"
    Print #ff, "set MsrParSize=16"
    Print #ff, "set LimitWinParSize="
    Print #ff, "set WinParSize="
    Print #ff, "set BypassDeviceEncryption="
    Print #ff, "set WinReParSize=2048"
    Print #ff, "set DiskId="
    Print #ff, "set UserLocale="
    Print #ff, "set ProductKey=VK7JG-NPHTM-C97JM-9MPGT-3V66T"
    Print #ff, "set Name="
    Print #ff, "set DisplayName="
    Print #ff, "set TimeZone=Central Standard Time"
    Print #ff, "set ComputerName="
    Print #ff, "set BypassWinRequirements="
    Print #ff, "set BypassQualityUpdatesDuringOobe="
    Print #ff, ""
    Print #ff, ":: We will ask the user if the answer file is being created for a BIOS based system or for a UEFI based system. For"
    Print #ff, ":: UEFI based systems, there are some unique settings that we need to ask for that don't apply to BIOS based systems."
    Print #ff, ":: If the user is creating an answer file for a UEFI based system then we need to gather that information. If not, then"
    Print #ff, ":: we can move on to gathering the information common to both BIOS and UEFI systems."
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Select the system type on which this answer file will be used:"
    Print #ff, "echo."
    Print #ff, "echo 1 - BIOS based system"
    Print #ff, "echo 2 - UEFI based system"
    Print #ff, "echo."
    Print #ff, "choice /C 12 /N /M "; Chr$(34); "Press 1 or 2 to make your selection:"; Chr$(34); ""
    Print #ff, "if errorlevel 2 set SystemType=UEFI & goto GetInfoForUefiSys"
    Print #ff, "if errorlevel 1 set SystemType=BIOS & goto CommonInfo"
    Print #ff, ""
    Print #ff, ":GetInfoForUefiSys"
    Print #ff, ""
    Print #ff, ":: The information gathered from the user in this section is needed only for UEFI based systems."
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Enter the size in MB to make the EFI partition. On most systems, a size of 100 MB is fine, but on 4k native format"
    Print #ff, "echo drives you should make this 260 MB since the minimum partition size for FAT32 on those drives is 260 MB."
    Print #ff, "echo TIP: If you want to guarantee compatibility on any system, use 260 MB. By doing so, you can use the same"
    Print #ff, "echo answer file on any system."
    Print #ff, "echo."
    Print #ff, "set /p EfiParSize="; Chr$(34); "Enter the size in MB to make the EFI partition or ENTER for 260 MB: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Enter the size for the Microsoft Reserved partition (MSR) in MB. It is suggested to use 16 MB."
    Print #ff, "echo."
    Print #ff, "set /p MsrParSize="; Chr$(34); "Enter the size in MB to make the Microsoft Reserved partition (MSR) or ENTER for 16 MB: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Do you want to prevent automatic device encryption on eligible systems?"
    Print #ff, "echo."
    Print #ff, "echo Windows can automatically encrypt the Windows partition on some UEFI systems. We can prevent this from happening."
    Print #ff, "echo."
    Print #ff, "choice /c YN /n /m "; Chr$(34); "Prevent automatic device encryption? Press Y for YES or N for NO:"; Chr$(34); ""
    Print #ff, "if errorlevel 2 set BypassDeviceEncryption=N & goto AskAboutLimitingSize"
    Print #ff, "if errorlevel 1 set BypassDeviceEncryption=Y & goto AskAboutLimitingSize"
    Print #ff, ""
    Print #ff, ":AskAboutLimitingSize"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Do you want to limit the size of the Windows partition? If you do not limit the size then the Windows partition will"
    Print #ff, "echo occupy all remaining space on the drive not used by the other partitions. If you choose to limit the size of the"
    Print #ff, "echo Windows partition, you will be asked what size to make that partition and another partition will be created that"
    Print #ff, "echo will occupy any space left on the drive. You can use that partition for anything you want."
    Print #ff, "echo."
    Print #ff, "choice /c YN /n /m "; Chr$(34); "Limit the size of the Windows partition? Press Y for YES or N for NO:"; Chr$(34); ""
    Print #ff, "if errorlevel 2 set LimitWinParSize=N & goto CommonInfo"
    Print #ff, "if errorlevel 1 set LimitWinParSize=Y & goto LimitSize"
    Print #ff, ""
    Print #ff, ":LimitSize"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "set /p WinParSize="; Chr$(34); "Enter the size in MB to make the Windows partition: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, ":: We will now begin to gather the information common to both BIOS and UEFI systems."
    Print #ff, ""
    Print #ff, ":CommonInfo"
    Print #ff, ""
    Print #ff, ":: The information gathered from the user in this section applies to both BIOS and UEFI based systems."
    Print #ff, ""
    Print #ff, ":GetWinReParSize"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Enter the size in MB to make the Windows Recovery Environment (WinRE) partition. You should make this a minimum of"
    Print #ff, "echo 750 MB but it is suggested to use 1000 MB if you can afford the space. I typically use 2000 MB because Microsoft"
    Print #ff, "echo seems to have been increasing the amount of space used in this partition lately."
    Print #ff, "echo."
    Print #ff, "set /p WinReParSize="; Chr$(34); "Enter the size in MB to make the WinRE partition or ENTER for 2048 MB: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, ":GetDiskId"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo On the next screen you will be asked to enter the Disk ID of the disk to which Windows should be installed."
    Print #ff, "echo."
    Print #ff, "echo IMPORTANT: The disk number that you specify here will be ERASED when Windows is installed. Do NOT use your running"
    Print #ff, "echo Windows installation to try to determine the disk ID because disk IDs during Windows setup may be different than while"
    Print #ff, "echo running Windows. If you have not already done so, you should follow these steps to determine the correct disk ID:"
    Print #ff, "echo."
    Print #ff, "echo 1) Create the Windows installation media that you will use to install Windows now. Do NOT include an autounattend.xml"
    Print #ff, "echo    answer file!"
    Print #ff, "echo."
    Print #ff, "echo 2) Boot from that media."
    Print #ff, "echo."
    Print #ff, "echo 3) At the very first static screen, press SHIFT + F10 to open a command prompt."
    Print #ff, "echo."
    Print #ff, "echo 4) At the command prompt, run "; Chr$(34); "diskpart"; Chr$(34); "."
    Print #ff, "echo."
    Print #ff, "echo 5) Once diskpart has started, run the command "; Chr$(34); "list disk"; Chr$(34); ". Note the disk ID (disk number) of the disk to which you"
    Print #ff, "echo    will install Windows. If the information shown is not enough to allow you to determine the correct disk, then"
    Print #ff, "echo    select a disk and show details for that disk to get more info. You can do this for as many disks as needed."
    Print #ff, "echo    EXAMPLE: "; Chr$(34); "select disk 0"; Chr$(34); ", then "; Chr$(34); "detail disk"; Chr$(34); "."
    Print #ff, "echo 6) Run "; Chr$(34); "exit"; Chr$(34); " twice to close diskpart and the command prompt."
    Print #ff, "echo."
    Print #ff, "echo 7) Reboot the system back into Windows."
    Print #ff, "echo."
    Print #ff, "echo IMPORTANT: Once the disk ID is determined, don't add or remove drives as the disk ID may then change!"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "pause"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "set /p DiskId="; Chr$(34); "Enter the disk ID of the disk to which Windows should be installed: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, ":BypassRequirements"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Do you want to bypass the Windows 11 system requirements check?"
    Print #ff, "echo."
    Print #ff, "echo Note that this option is safe to use even on systems that do meet Windows 11 requirements, it simply will have no"
    Print #ff, "echo effect on those systems. By using this option, you can use the same answer on both systems that meet requirements"
    Print #ff, "echo and on systems that do not meet requirements."
    Print #ff, "echo."
    Print #ff, "choice /C YN /N /M "; Chr$(34); "Press Y for YES or N for NO:"; Chr$(34); ""
    Print #ff, "if errorlevel 2 set BypassWinRequirements=N & goto QualityUpdates"
    Print #ff, "if errorlevel 1 set BypassWinRequirements=Y & goto QualityUpdates"
    Print #ff, ""
    Print #ff, ":QualityUpdates"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Do you want to prevent checks for quality updates during Windows installation?"
    Print #ff, "echo."
    Print #ff, "echo Windows setup can now check for quality updates during Windows setup. This can considerably lengthen the amount of"
    Print #ff, "echo time that it takes to install Windows. If you select YES then we will prevent these checks during setup."
    Print #ff, "echo."
    Print #ff, "choice /C YN /N /M "; Chr$(34); "Press Y for YES or N for NO:"; Chr$(34); ""
    Print #ff, "if errorlevel 2 set BypassQualityUpdatesDuringOobe=N & goto GetProdKey"
    Print #ff, "if errorlevel 1 set BypassQualityUpdatesDuringOobe=Y & goto GetProdKey"
    Print #ff, ""
    Print #ff, ":GetProdKey"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo We need the generic product key for the edition of Windows that you wish to install. Below are the most commonly"
    Print #ff, "echo used product keys. For other keys, please visit this link:"
    Print #ff, "echo."
    Print #ff, "echo https://www.elevenforum.com/t/generic-product-keys-to-install-or-upgrade-windows-11-editions.3713/"
    Print #ff, "echo."
    Print #ff, "echo Windows 10 or 11 Home Single Language:  BT79Q-G7N6G-PGBYW-4YWX6-6F4BT"
    Print #ff, "echo Windows 10 or 11 Home:                  YTMG3-N6DKC-DKB77-7M9GH-8HVX7"
    Print #ff, "echo Windows 10 or 11 Pro:                   VK7JG-NPHTM-C97JM-9MPGT-3V66T"
    Print #ff, "echo."
    Print #ff, "set /p ProductKey="; Chr$(34); "Enter the Windows product key or press ENTER to use the Win Pro edition key: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, ":GetUserLocale"
    Print #ff, "cls"
    Print #ff, "echo Enter the User Locale to be used. Normally you should select "; Chr$(34); "en-US"; Chr$(34); ". However, you can instead opt to use "; Chr$(34); "en-001"; Chr$(34); "."
    Print #ff, "echo If you use en-001, you will end up with a very clean Start screen but there are steps that you MUST take after"
    Print #ff, "echo Windows installation has finished. If you use en-001, then perform these steps immediately after Windows installation:"
    Print #ff, "echo."
    Print #ff, "echo 1) Enable networking so that Windows has access to the Internet. Look at the Start screen. If you see any"
    Print #ff, "echo    "; Chr$(34); "placeholders"; Chr$(34); ", that is greyed out areas in places where icons could be placed, then the system has not yet been"
    Print #ff, "echo    able to access the Internet. Correct this now before moving to step 2. Once the placeholders have disappeared"
    Print #ff, "echo    you can move on to step 2."
    Print #ff, "echo."
    Print #ff, "echo 2) Open Settings - Time ^& language - Language ^& region. Change "; Chr$(34); "Country or region"; Chr$(34); " to "; Chr$(34); "United States"; Chr$(34); " and change"
    Print #ff, "echo    "; Chr$(34); "Regional format"; Chr$(34); " to "; Chr$(34); "English (United States)"; Chr$(34); "."
    Print #ff, "echo."
    Print #ff, "choice /C 12 /N /M "; Chr$(34); "Press 1 for en-US or 2 for en-001:"; Chr$(34); ""
    Print #ff, "if errorlevel 2 set UserLocale=en-001 & goto GetName"
    Print #ff, "if errorlevel 1 set UserLocale=en-US & goto GetName"
    Print #ff, ""
    Print #ff, ":GetName"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Enter the user name that you want to use. Note that this is the user name that will be used to create a local"
    Print #ff, "echo Administrator account for you. On the next screen you will be asked for the "; Chr$(34); "Display Name"; Chr$(34); ". The display name"
    Print #ff, "echo is a friendly name or full name of the user that is used in places like the lock screen."
    Print #ff, "echo."
    Print #ff, "echo EXAMPLE: You might specify "; Chr$(34); "WinUser"; Chr$(34); " as the User Name and "; Chr$(34); "Windows User"; Chr$(34); " as the Display Name."
    Print #ff, "echo."
    Print #ff, "set /p Name="; Chr$(34); "Enter the User Name to be created as a local Administrator: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, ":GetDisplayName"
    Print #ff, "cls"
    Print #ff, "set /p DisplayName="; Chr$(34); "Enter the User Display Name / Full User Name: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, ":GetTimezone"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Please enter the time zone that this computer will be in. For example, "; Chr$(34); "Central Standard Time"; Chr$(34); ". To get a list of valid"
    Print #ff, "echo time zones, run the command "; Chr$(34); "tzutil /L"; Chr$(34); ". The second line of each group is the name that you can specify here."
    Print #ff, "echo."
    Print #ff, "set /p TimeZone="; Chr$(34); "Enter timezone or press ENTER to use Central Standard Time: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, ":GetComputerName"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Enter the computer name that you want to use or leave blank. If you leave this blank (by just pressing ENTER), then a"
    Print #ff, "echo random name will be assigned and you can change the name after installation. It is recommended to leave this blank to"
    Print #ff, "echo avoid accidentally assigning the same name to multiple machines if you use the same answer file for multiple machines."
    Print #ff, "echo."
    Print #ff, "set /p ComputerName="; Chr$(34); "Enter the computer name: "; Chr$(34); ""
    Print #ff, ""
    Print #ff, ":: We are done gathering all needed information through interactive mode. We will now"
    Print #ff, ":: continue as if we were running in automatic mode."
    Print #ff, ""
    Print #ff, ":RunAutoMode"
    Print #ff, ""
    Print #ff, "::::::::::::::::::::::::::::::::::::::::"
    Print #ff, ":: Answer File Generation Begins Here ::"
    Print #ff, "::::::::::::::::::::::::::::::::::::::::"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Generating answer file. Please standby..."
    Print #ff, ""
    Print #ff, ":: Include this section for both BIOS and UEFI systems:"
    Print #ff, ""
    Print #ff, "echo ^<?xml version="; Chr$(34); "1.0"; Chr$(34); " encoding="; Chr$(34); "utf-8"; Chr$(34); "?^> > autounattend.xml"
    Print #ff, "echo ^<unattend xmlns="; Chr$(34); "urn:schemas-microsoft-com:unattend"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo    ^<settings pass="; Chr$(34); "windowsPE"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo        ^<component name="; Chr$(34); "Microsoft-Windows-International-Core-WinPE"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo            ^<SetupUILanguage^> >> autounattend.xml"
    Print #ff, "echo                ^<UILanguage^>en-US^</UILanguage^> >> autounattend.xml"
    Print #ff, "echo            ^</SetupUILanguage^> >> autounattend.xml"
    Print #ff, "echo            ^<InputLocale^>en-US^</InputLocale^> >> autounattend.xml"
    Print #ff, "echo            ^<SystemLocale^>en-US^</SystemLocale^> >> autounattend.xml"
    Print #ff, "echo            ^<UILanguage^>en-US^</UILanguage^> >> autounattend.xml"
    Print #ff, "echo            ^<UserLocale^>%UserLocale%^</UserLocale^> >> autounattend.xml"
    Print #ff, "echo        ^</component^> >> autounattend.xml"
    Print #ff, "echo        ^<component name="; Chr$(34); "Microsoft-Windows-Setup"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo            ^<ImageInstall^> >> autounattend.xml"
    Print #ff, "echo                ^<OSImage^> >> autounattend.xml"
    Print #ff, "echo                    ^<InstallTo^> >> autounattend.xml"
    Print #ff, "echo                        ^<DiskID^>%DiskId%^</DiskID^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: If Windows is being installed on a UEFI system, then set the partition to which Windows should be installed to 3. On a"
    Print #ff, ":: BIOS based system, use partition 2."
    Print #ff, ""
    Print #ff, "if %SystemType%==UEFI ("
    Print #ff, "set InstallPar=3"
    Print #ff, ") else ("
    Print #ff, "set InstallPar=2"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "echo                        ^<PartitionID^>%InstallPar%^</PartitionID^> >> autounattend.xml"
    Print #ff, "echo                    ^</InstallTo^> >> autounattend.xml"
    Print #ff, "echo                    ^<Compact^>false^</Compact^> >> autounattend.xml"
    Print #ff, "echo                ^</OSImage^> >> autounattend.xml"
    Print #ff, "echo            ^</ImageInstall^> >> autounattend.xml"
    Print #ff, "echo            ^<UserData^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Use product key for version to be installed"
    Print #ff, ""
    Print #ff, "echo                ^<ProductKey^> >> autounattend.xml"
    Print #ff, "echo                    ^<Key^>%ProductKey%^</Key^> >> autounattend.xml"
    Print #ff, "echo                ^</ProductKey^> >> autounattend.xml"
    Print #ff, "echo                ^<AcceptEula^>true^</AcceptEula^> >> autounattend.xml"
    Print #ff, "echo            ^</UserData^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Create the RunSynchronous block to allow us to run specific commands"
    Print #ff, ""
    Print #ff, ":: NOTE: The "; Chr$(34); "Order"; Chr$(34); " specified for commands run here must be in order (1, 2, 3, etc.). Skipping numbers is not"
    Print #ff, ":: permitted. To facilitate this, we will set a counter with an initial value of 1. After each command we increment"
    Print #ff, ":: the counter by 1."
    Print #ff, ""
    Print #ff, "set /A Phase1CommandCounter=1"
    Print #ff, ""
    Print #ff, "echo            ^<RunSynchronous^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "If %BypassWinRequirements%==N goto BypassRequirementsDone"
    Print #ff, ""
    Print #ff, ":: These commands bypass Win 11 requirements"
    Print #ff, ""
    Print #ff, "echo                ^<RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<Order^>%Phase1CommandCounter%^</Order^> >> autounattend.xml"
    Print #ff, "echo                    ^<Path^>reg add HKLM\System\Setup\LabConfig /v BypassTPMCheck /t reg_dword /d 0x00000001 /f^</Path^> >> autounattend.xml"
    Print #ff, "echo                ^</RunSynchronousCommand^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "set /A Phase1CommandCounter +=1"
    Print #ff, ""
    Print #ff, "echo                ^<RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<Order^>%Phase1CommandCounter%^</Order^> >> autounattend.xml"
    Print #ff, "echo                    ^<Path^>reg add HKLM\System\Setup\LabConfig /v BypassSecureBootCheck /t reg_dword /d 0x00000001 /f^</Path^> >> autounattend.xml"
    Print #ff, "echo                ^</RunSynchronousCommand^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "set /A Phase1CommandCounter +=1"
    Print #ff, ""
    Print #ff, "echo                ^<RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<Order^>%Phase1CommandCounter%^</Order^> >> autounattend.xml"
    Print #ff, "echo                    ^<Path^>reg add HKLM\System\Setup\LabConfig /v BypassRAMCheck /t reg_dword /d 0x00000001 /f^</Path^> >> autounattend.xml"
    Print #ff, "echo                ^</RunSynchronousCommand^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "set /A Phase1CommandCounter +=1"
    Print #ff, ""
    Print #ff, ":BypassRequirementsDone"
    Print #ff, ""
    Print #ff, ":: Do these steps for UEFI systems"
    Print #ff, ""
    Print #ff, ":: This command performs all the disk setup operations for UEFI based systems only if user wants to limit the size of"
    Print #ff, ":: the Windows partition. We skip this for BIOS based systems."
    Print #ff, ""
    Print #ff, "if %SystemType%==BIOS goto BiosPartitioning"
    Print #ff, "if %LimitWinParSize%==N goto CreateFullSizeWinPar "
    Print #ff, ""
    Print #ff, "echo                ^<RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<Order^>%Phase1CommandCounter%^</Order^> >> autounattend.xml"
    Print #ff, "echo                    ^<Path^>cmd /c (for %%a in ("; Chr$(34); "sel dis %DiskId%"; Chr$(34); " "; Chr$(34); "cle"; Chr$(34); " "; Chr$(34); "con gpt"; Chr$(34); " "; Chr$(34); "cre par efi size=%EfiParSize%"; Chr$(34); " "; Chr$(34); "for quick fs=fat32"; Chr$(34); " "; Chr$(34); "cre par msr size=%MsrParSize%"; Chr$(34); " "; Chr$(34); "cre par pri size=%WinParSize%"; Chr$(34); " "; Chr$(34); "format quick fs=ntfs label="; Chr$(34); "Windows"; Chr$(34); ""; Chr$(34); " "; Chr$(34); "cre par pri size=%WinReParSize%"; Chr$(34); " "; Chr$(34); "for quick fs=ntfs"; Chr$(34); " "; Chr$(34); "set id=de94bba4-06d1-4d40-a16a-bfd50179d6ac"; Chr$(34); " "; Chr$(34); "gpt attributes=0x8000000000000001"; Chr$(34); " "; Chr$(34); "create partition primary"; Chr$(34); " "; Chr$(34); "format quick fs=ntfs"; Chr$(34); ") do @echo %%~a) ^&gt; X:\UEFI.txt ^&amp; diskpart /s X:\UEFI.txt^</Path^> >> autounattend.xml"
    Print #ff, "echo                ^</RunSynchronousCommand^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "set /A Phase1CommandCounter +=1"
    Print #ff, ""
    Print #ff, ":: This was the last command to be run so we can close out the RunSynchronous block"
    Print #ff, ""
    Print #ff, "echo            ^</RunSynchronous^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "goto DonePartitioning"
    Print #ff, ""
    Print #ff, ":CreateFullSizeWinPar"
    Print #ff, ""
    Print #ff, "echo                ^<RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<Order^>%Phase1CommandCounter%^</Order^> >> autounattend.xml"
    Print #ff, "echo                    ^<Path^>cmd /c (for %%a in ("; Chr$(34); "sel dis %DiskId%"; Chr$(34); " "; Chr$(34); "cle"; Chr$(34); " "; Chr$(34); "con gpt"; Chr$(34); " "; Chr$(34); "cre par efi size=%EfiParSize%"; Chr$(34); " "; Chr$(34); "for quick fs=fat32"; Chr$(34); " "; Chr$(34); "cre par msr size=%MsrParSize%"; Chr$(34); " "; Chr$(34); "cre par pri"; Chr$(34); " "; Chr$(34); "shr minimum=%WinReParSize%"; Chr$(34); " "; Chr$(34); "for quick fs=ntfs label="; Chr$(34); "Windows"; Chr$(34); ""; Chr$(34); " "; Chr$(34); "cre par pri"; Chr$(34); " "; Chr$(34); "for quick fs=ntfs"; Chr$(34); " "; Chr$(34); "set id=de94bba4-06d1-4d40-a16a-bfd50179d6ac"; Chr$(34); " "; Chr$(34); "gpt attributes=0x8000000000000001"; Chr$(34); ") do @echo %%~a) ^&gt; X:\UEFI.txt ^&amp; diskpart /s X:\UEFI.txt^</Path^> >> autounattend.xml"
    Print #ff, "echo                ^</RunSynchronousCommand^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "set /A Phase1CommandCounter +=1"
    Print #ff, ""
    Print #ff, ":: This was the last command to be run so we can close out the RunSynchronous block"
    Print #ff, ""
    Print #ff, "echo            ^</RunSynchronous^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "goto DonePartitioning"
    Print #ff, ""
    Print #ff, ":BiosPartitioning"
    Print #ff, ""
    Print #ff, ":: This was the last command to be run so we can close out the RunSynchronous block"
    Print #ff, ""
    Print #ff, "echo            ^</RunSynchronous^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: These are the operations needed to setup the drive for a BIOS based system."
    Print #ff, ""
    Print #ff, "echo            ^<DiskConfiguration^> >> autounattend.xml"
    Print #ff, "echo                ^<Disk wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<CreatePartitions^> >> autounattend.xml"
    Print #ff, "echo                        ^<CreatePartition wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                            ^<Order^>1^</Order^> >> autounattend.xml"
    Print #ff, "echo                            ^<Size^>%WinReParSize%^</Size^> >> autounattend.xml"
    Print #ff, "echo                            ^<Type^>Primary^</Type^> >> autounattend.xml"
    Print #ff, "echo                        ^</CreatePartition^> >> autounattend.xml"
    Print #ff, "echo                        ^<CreatePartition wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                            ^<Extend^>true^</Extend^> >> autounattend.xml"
    Print #ff, "echo                            ^<Order^>2^</Order^> >> autounattend.xml"
    Print #ff, "echo                            ^<Type^>Primary^</Type^> >> autounattend.xml"
    Print #ff, "echo                        ^</CreatePartition^> >> autounattend.xml"
    Print #ff, "echo                    ^</CreatePartitions^> >> autounattend.xml"
    Print #ff, "echo                    ^<ModifyPartitions^> >> autounattend.xml"
    Print #ff, "echo                        ^<ModifyPartition wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                            ^<Active^>true^</Active^> >> autounattend.xml"
    Print #ff, "echo                            ^<Format^>NTFS^</Format^> >> autounattend.xml"
    Print #ff, "REM echo                            ^<Label^>System^</Label^> >> autounattend.xml"
    Print #ff, "echo                            ^<Order^>1^</Order^> >> autounattend.xml"
    Print #ff, "echo                            ^<PartitionID^>1^</PartitionID^> >> autounattend.xml"
    Print #ff, "echo                        ^</ModifyPartition^> >> autounattend.xml"
    Print #ff, "echo                        ^<ModifyPartition wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                            ^<Format^>NTFS^</Format^> >> autounattend.xml"
    Print #ff, "echo                            ^<Label^>Windows^</Label^> >> autounattend.xml"
    Print #ff, "echo                            ^<Letter^>C^</Letter^> >> autounattend.xml"
    Print #ff, "echo                            ^<Order^>2^</Order^> >> autounattend.xml"
    Print #ff, "echo                            ^<PartitionID^>2^</PartitionID^> >> autounattend.xml"
    Print #ff, "echo                        ^</ModifyPartition^> >> autounattend.xml"
    Print #ff, "echo                    ^</ModifyPartitions^> >> autounattend.xml"
    Print #ff, "echo                    ^<DiskID^>0^</DiskID^> >> autounattend.xml"
    Print #ff, "echo                    ^<WillWipeDisk^>true^</WillWipeDisk^> >> autounattend.xml"
    Print #ff, "echo                ^</Disk^> >> autounattend.xml"
    Print #ff, "echo            ^</DiskConfiguration^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":DonePartitioning"
    Print #ff, ""
    Print #ff, "echo        ^</component^> >> autounattend.xml"
    Print #ff, "echo    ^</settings^> >> autounattend.xml"
    Print #ff, "echo    ^<settings pass="; Chr$(34); "oobeSystem"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo        ^<component name="; Chr$(34); "Microsoft-Windows-International-Core"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo            ^<InputLocale^>en-US^</InputLocale^> >> autounattend.xml"
    Print #ff, "echo            ^<SystemLocale^>en-US^</SystemLocale^> >> autounattend.xml"
    Print #ff, "echo            ^<UILanguage^>en-US^</UILanguage^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "REM Set the proper UserLocale setting:"
    Print #ff, ""
    Print #ff, "echo            ^<UserLocale^>%UserLocale%^</UserLocale^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "echo        ^</component^> >> autounattend.xml"
    Print #ff, "echo        ^<component name="; Chr$(34); "Microsoft-Windows-Shell-Setup"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo            ^<OOBE^> >> autounattend.xml"
    Print #ff, "echo                ^<HideEULAPage^>true^</HideEULAPage^> >> autounattend.xml"
    Print #ff, "echo                ^<HideOEMRegistrationScreen^>true^</HideOEMRegistrationScreen^> >> autounattend.xml"
    Print #ff, "echo                ^<HideOnlineAccountScreens^>true^</HideOnlineAccountScreens^> >> autounattend.xml"
    Print #ff, "echo                ^<HideWirelessSetupInOOBE^>true^</HideWirelessSetupInOOBE^> >> autounattend.xml"
    Print #ff, "echo                ^<ProtectYourPC^>1^</ProtectYourPC^> >> autounattend.xml"
    Print #ff, "echo                ^<UnattendEnableRetailDemo^>false^</UnattendEnableRetailDemo^> >> autounattend.xml"
    Print #ff, "echo            ^</OOBE^> >> autounattend.xml"
    Print #ff, "echo            ^<UserAccounts^> >> autounattend.xml"
    Print #ff, "echo                ^<LocalAccounts^> >> autounattend.xml"
    Print #ff, "echo                    ^<LocalAccount wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Using a password of "; Chr$(34); "Password1"; Chr$(34); ""
    Print #ff, ""
    Print #ff, "echo                        ^<Password^> >> autounattend.xml"
    Print #ff, "echo                            ^<Value^>Password1^</Value^> >> autounattend.xml"
    Print #ff, "echo                            ^<PlainText^>true^</PlainText^> >> autounattend.xml"
    Print #ff, "echo                        ^</Password^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Create the local user account and make a part of the Administrators group."
    Print #ff, ""
    Print #ff, "echo                        ^<DisplayName^>%DisplayName%^</DisplayName^> >> autounattend.xml"
    Print #ff, "echo                        ^<Group^>Administrators^</Group^> >> autounattend.xml"
    Print #ff, "echo                        ^<Name^>%Name%^</Name^> >> autounattend.xml"
    Print #ff, "echo                    ^</LocalAccount^> >> autounattend.xml"
    Print #ff, "echo                ^</LocalAccounts^> >> autounattend.xml"
    Print #ff, "echo            ^</UserAccounts^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Set the time zone"
    Print #ff, ""
    Print #ff, "echo            ^<TimeZone^>%TimeZone%^</TimeZone^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Add a registry entry to resolve a bug related to autologon. This answer file will autologon just one time in order"
    Print #ff, ":: to complete Windows setup. Later in this answer file you will see where we specify a logon count of 1. The bug is"
    Print #ff, ":: that Windows will autologon one time more than specified. So, you would think that you could specify zero and that"
    Print #ff, ":: this would result in one logon. Unfortunately, the system does properly understand that zero means nerver logon. The"
    Print #ff, ":: registry entry works around this bug."
    Print #ff, ""
    Print #ff, ":: Set a counter called FirstLogonCommandCounter to keep track of the order of commands. We will set the initial value"
    Print #ff, ":: to 1 and increment each time a command is added."
    Print #ff, ""
    Print #ff, "set /A FirstLogonCommandCounter=1"
    Print #ff, ""
    Print #ff, "echo            ^<FirstLogonCommands^> >> autounattend.xml"
    Print #ff, "echo                ^<SynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<CommandLine^>reg add ^&quot;HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon^&quot; /v AutoLogonCount /t REG_DWORD /d 0 /f^</CommandLine^> >> autounattend.xml"
    Print #ff, "echo                    ^<Order^>%FirstLogonCommandCounter%^</Order^> >> autounattend.xml"
    Print #ff, "echo                ^</SynchronousCommand^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "set /A FirstLogonCommandCounter+=1"
    Print #ff, ""
    Print #ff, ":: Use below command only if bypassing quality updates during setup"
    Print #ff, ""
    Print #ff, "if %BypassQualityUpdatesDuringOobe%==N goto DoneWithQualityUpdates"
    Print #ff, ""
    Print #ff, "echo                ^<SynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<Order^>%FirstLogonCommandCounter%^</Order^> >> autounattend.xml"
    Print #ff, "echo                    ^<CommandLine^>powershell.exe -Command ^&quot;Get-NetAdapter ^| ForEach-Object { Enable-NetAdapter -Name $_.Name -Confirm:$false }^&quot;^</CommandLine^> >> autounattend.xml"
    Print #ff, "echo                ^</SynchronousCommand^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "set /A %FirstLogonCommandCounter%+=1"
    Print #ff, ""
    Print #ff, ":DoneWithQualityUpdates"
    Print #ff, ""
    Print #ff, ":: Setup is not fully completed until the user logs on for the first time. We are setting a one-time automatic logon."
    Print #ff, ""
    Print #ff, "echo            ^</FirstLogonCommands^> >> autounattend.xml"
    Print #ff, "echo            ^<AutoLogon^> >> autounattend.xml"
    Print #ff, "echo                ^<Password^> >> autounattend.xml"
    Print #ff, "echo                    ^<Value^>Password1^</Value^> >> autounattend.xml"
    Print #ff, "echo                    ^<PlainText^>true^</PlainText^> >> autounattend.xml"
    Print #ff, "echo                ^</Password^> >> autounattend.xml"
    Print #ff, "echo                ^<Username^>%Name%^</Username^> >> autounattend.xml"
    Print #ff, "echo                ^<Enabled^>true^</Enabled^> >> autounattend.xml"
    Print #ff, "echo                ^<LogonCount^>1^</LogonCount^> >> autounattend.xml"
    Print #ff, "echo            ^</AutoLogon^> >> autounattend.xml"
    Print #ff, "echo        ^</component^> >> autounattend.xml"
    Print #ff, "echo    ^</settings^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Start the Specialize pass"
    Print #ff, ""
    Print #ff, "echo    ^<settings pass="; Chr$(34); "specialize"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo        ^<component name="; Chr$(34); "Microsoft-Windows-Shell-Setup"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Set the computer name. If name was not provided, then a random name is assigned."
    Print #ff, ""
    Print #ff, "echo            ^<ComputerName^>%ComputerName%^</ComputerName^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Set the time zone"
    Print #ff, ""
    Print #ff, "echo            ^<TimeZone^>%TimeZone%^</TimeZone^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Create the Windows Deployment section"
    Print #ff, ""
    Print #ff, "echo        ^</component^> >> autounattend.xml"
    Print #ff, "echo        ^<component name="; Chr$(34); "Microsoft-Windows-Deployment"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Use a command to bypass quality updates during setup. Use another command to prevent auto device encryption."
    Print #ff, ":: First, we create the RunSynchronous block so that we can add those commands."
    Print #ff, ""
    Print #ff, "echo            ^<RunSynchronous^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Create a counter called Phase4CommandCounter to keep trackof command numbers. Increment every"
    Print #ff, ":: time a command is added."
    Print #ff, ""
    Print #ff, "set /A Phase4CommandCounter=1"
    Print #ff, ""
    Print #ff, ":: This command disables networking so that quality updates cannot be installed during setup. As a result, add it only"
    Print #ff, ":: if quality updates are being bypassed."
    Print #ff, ""
    Print #ff, ":: For the "; Chr$(34); "if"; Chr$(34); " statements below, the order is important. If system type is specified as BIOS but we check the 2nd if"
    Print #ff, ":: statement first, this would cause an error because the variable that we are checking is empty causing an error."
    Print #ff, ""
    Print #ff, "if %SystemType%==BIOS goto DeviceEncryption"
    Print #ff, "if %BypassQualityUpdatesDuringOobe%==N goto DeviceEncryption"
    Print #ff, ""
    Print #ff, "echo                ^<RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<Order^>%Phase4CommandCounter%^</Order^> >> autounattend.xml"
    Print #ff, "echo                    ^<Path^>powershell.exe -Command ^&quot;Get-NetAdapter ^| ForEach-Object { Disable-NetAdapter -Name $_.Name -Confirm:$false }^&quot;^</Path^> >> autounattend.xml"
    Print #ff, "echo                ^</RunSynchronousCommand^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "set /A Phase4CommandCounter+=1"
    Print #ff, ""
    Print #ff, ":DeviceEncryption"
    Print #ff, ""
    Print #ff, ":: This command bypasses automatic device encryption. Add it only if user elected to bypass device encryption."
    Print #ff, ""
    Print #ff, ":: For the "; Chr$(34); "if"; Chr$(34); " statements below, the order is important. If system type is specified as BIOS but we check the 2nd if"
    Print #ff, ":: statement first, this would cause an error because the variable that we are checking is empty causing an error."
    Print #ff, ""
    Print #ff, "if %SystemType%==BIOS goto DevEncryptDone"
    Print #ff, "if %BypassDeviceEncryption%==N goto DevEncryptDone"
    Print #ff, ""
    Print #ff, "echo                ^<RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); "^> >> autounattend.xml"
    Print #ff, "echo                    ^<Order^>%Phase4CommandCounter%^</Order^> >> autounattend.xml"
    Print #ff, "echo                    ^<Path^>reg add HKLM\System\CurrentControlSet\Control\BitLocker /v PreventDeviceEncryption /t reg_dword /d 0x00000001 /f^</Path^> >> autounattend.xml"
    Print #ff, "echo                ^</RunSynchronousCommand^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, "set /A Phase4CommandCounter+=1"
    Print #ff, ""
    Print #ff, ":DevEncryptDone"
    Print #ff, ""
    Print #ff, ":: Close the RunSynchronous block"
    Print #ff, ""
    Print #ff, "echo            ^</RunSynchronous^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":: Close everything else."
    Print #ff, ""
    Print #ff, "echo        ^</component^> >> autounattend.xml"
    Print #ff, "echo    ^</settings^> >> autounattend.xml"
    Print #ff, "echo ^</unattend^> >> autounattend.xml"
    Print #ff, ""
    Print #ff, ":DoneGenerating"
    Print #ff, ""
    Print #ff, "move /Y "; Chr$(34); "autounattend.xml"; Chr$(34); " %Vol1%:\ > NUL"
    Print #ff, "set af_FileNameOnly=On the fly user generated answer file"
    Print #ff, "popd"
    Print #ff, "cls"
    Print #ff, "echo Answer file generation completed."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, ""
    Print #ff, "::::::::::::::::::::::::::::::::::"
    Print #ff, ":: END OF ANSWER FILE GENERATOR ::"
    Print #ff, "::::::::::::::::::::::::::::::::::"
    Print #ff, ""
    Print #ff, ":AnsGenDone"
    Print #ff, ""
    Print #ff, "REM Create a backup of the files on the first volume for later recovery."
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: Creating a backup copy of the current configuration ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "robocopy %Vol1%:\ %Vol2%:\PE_Backup /mir /xf autounattend.xml /xd "; Chr$(34); "system volume information"; Chr$(34); " $recycle.bin > NUL"
    Print #ff, ""
    Print #ff, ":CopyFiles"
    Print #ff, ""
    Print #ff, "REM Mount the ISO image and copy the files to the drive. We only need the sources folder on the second volume."
    Print #ff, "REM  On the first volume, we want everything else. We also want the file called BOOT.WIM in the sources"
    Print #ff, "REM  folder on the first volume."
    Print #ff, ""
    Print #ff, "echo :::::::::::::::::::::::::::::"
    Print #ff, "echo :: Mounting selected image ::"
    Print #ff, "echo :::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, "echo The image to be deployed is:"
    Print #ff, "echo     %FileNameOnly%"
    Print #ff, "echo."
    Print #ff, "if NOT "; Chr$(34); "%af_FileNameOnly%"; Chr$(34); "=="; Chr$(34); Chr$(34); " ("
    Print #ff, "echo You chose to add this answer file:"
    Print #ff, "echo     %af_FileNameOnly%"
    Print #ff, "echo."
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "REM Mount the image."
    Print #ff, ""
    Print #ff, "powershell.exe -command "; Chr$(34); "Mount-DiskImage "; Chr$(34); "'%SourceISOImage%'"; Chr$(34); " -PassThru | Get-Volume"; Chr$(34); " > MountInfo.txt"
    Print #ff, ""
    Print #ff, "REM Get drive letter (includes the colon)."
    Print #ff, "REM We need to skip 3 lines in order to read the 4th line of text in the MountInfo.txt file."
    Print #ff, ""
    Print #ff, "for /F "; Chr$(34); "skip=3 delims="; Chr$(34); " %%a in (MountInfo.txt) do ("
    Print #ff, "   set LineContents=%%a"
    Print #ff, "   goto Evaluate"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, ":Evaluate"
    Print #ff, ""
    Print #ff, "REM We are done with MountInfo.txt. Delete it."
    Print #ff, ""
    Print #ff, "del MountInfo.txt"
    Print #ff, "set LineContents=%LineContents:~0,1%"
    Print #ff, "set MountedImageDriveLetter=%LineContents%:"
    Print #ff, ""
    Print #ff, "REM Check to see if the root of the source image has an autounattend.xml answer file AND if the user selected an answer"
    Print #ff, "REM file from the list of available answer files. If so, we need the user to clarify which to use."
    Print #ff, "REM Start by setting default values:"
    Print #ff, ""
    Print #ff, "set SourceAF=N"
    Print #ff, "set UserAF=N"
    Print #ff, ""
    Print #ff, "if exist "; Chr$(34); "%MountedImageDriveLetter%\autounattend.xml"; Chr$(34); " ("
    Print #ff, "set SourceAF=Y"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if NOT "; Chr$(34); "%af_UserAnswerFile%"; Chr$(34); "=="; Chr$(34); Chr$(34); " ("
    Print #ff, "set UserAF=Y"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if %SourceAF%==Y if %UserAF%==Y goto SelectAnswerFileSource"
    Print #ff, ""
    Print #ff, "goto StartFileCopy"
    Print #ff, ""
    Print #ff, ":SelectAnswerFileSource"
    Print #ff, ""
    Print #ff, "REM We arrive here if the source has an answer file AND user selected an answer file"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo CAUTION: The image that you selected already has an answer file AND you have selected another answer file to use."
    Print #ff, "echo          Please specify whether you want to use the answer file from the image or the one you selected."
    Print #ff, "echo."
    Print #ff, "echo Select 1 to use the answer file from the image."
    Print #ff, "echo Select 2 to use the answer file that you selected from the list."
    Print #ff, "echo."
    Print #ff, "choice /C 12 /N /M "; Chr$(34); "Press 1 or 2 to make your choice: "; Chr$(34)
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "if errorlevel==2 ("
    Print #ff, "set SourceAF=N"
    Print #ff, "set UserAF=Y"
    Print #ff, "goto DoneSelectingAF"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if errorlevel==1 ("
    Print #ff, "set SourceAF=Y"
    Print #ff, "set UserAF=N"
    Print #ff, "goto DoneSelectingAF"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, ":DoneSelectingAF"
    Print #ff, ""
    Print #ff, ":StartFileCopy"
    Print #ff, ""
    Print #ff, "REM Start the file copies"
    Print #ff, ""
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: Copying files. This may take a considerable amount  ::"
    Print #ff, "echo :: of time, especially if your disk is slow. Please be ::"
    Print #ff, "echo :: patient and allow the process to complete.          ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "REM :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "REM :: Check to see if the selected image is a single or dual architecture Windows image or if it is ::"
    Print #ff, "REM :: a Windows PE based media. If file named \sources\install.wim or \sources\install.esd exists,  ::"
    Print #ff, "REM :: then we have a single architecture image. If not, check for the same files in a \x64\sources  ::"
    Print #ff, "REM :: folder. That would indicate a dual architecture image. Finally, if none of those files exist  ::"
    Print #ff, "REM :: but a \sources\boot.wim file exists, then we have a Windows PE based image.                   ::"
    Print #ff, "REM ::                                                                                               ::"
    Print #ff, "REM :: For Windows images (not WinPE images), we will also add an ei.cfg file to the \sources        ::"
    Print #ff, "REM :: folder(s). The ei.cfg file will cause Windows setup to allow a user to select any edition of  ::"
    Print #ff, "REM :: Windows available in the image file. Without this file, if you are installing on a system     ::"
    Print #ff, "REM :: that shipped with Windows, setup will automatically start installing the same version of      ::"
    Print #ff, "REM :: Windows that the system originally shipped with. This is a problem if you have upgraded from  ::"
    Print #ff, "REM :: one edition to another, for example, from Home to Pro.                                        ::"
    Print #ff, "REM :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, ""
    Print #ff, "if exist %MountedImageDriveLetter%\sources\install.wim ("
    Print #ff, "set ImageType=Single"
    Print #ff, "goto CopySingle"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if exist %MountedImageDriveLetter%\sources\install.esd ("
    Print #ff, "set ImageType=Single"
    Print #ff, "goto CopySingle"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if exist %MountedImageDriveLetter%\x64\sources\install.wim ("
    Print #ff, "set ImageType=Dual"
    Print #ff, "goto CopyDual"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if exist %MountedImageDriveLetter%\x64\sources\install.esd ("
    Print #ff, "set ImageType=Dual"
    Print #ff, "goto CopyDual"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if exist %MountedImageDriveLetter%\sources\boot.wim ("
    Print #ff, "set ImageType=PE"
    Print #ff, "goto CopyPE"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "REM We reach this section only if no valid image was found."
    Print #ff, "REM We inform the user and delete the Windows PE backup that we created."
    Print #ff, ""
    Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'%SourceISOImage%'"; Chr$(34); ""; Chr$(34); " > NUL"
    Print #ff, "cls"
    Print #ff, "echo The file specified does not appear to be a valid image."
    Print #ff, "pause"
    Print #ff, "RD /Q /S %Vol2%:\PE_Backup > NUL 2>&1"
    Print #ff, "goto End"
    Print #ff, ""
    Print #ff, "REM the :CopySingle, :CopyDual, and :CopyPE sections copy the files needed from the seletced image to Vol1 and Vol2."
    Print #ff, ""
    Print #ff, ":CopySingle"
    Print #ff, ""
    Print #ff, "if %SourceAF% == Y ("
    Print #ff, "robocopy %MountedImageDriveLetter% %Vol1%:\ /mir /xf VOL1_M_MEDIA.WIM /xd sources "; Chr$(34); "system volume information"; Chr$(34); " $recycle.bin /256 /r:0 > NUL"
    Print #ff, ") ELSE ("
    Print #ff, "robocopy %MountedImageDriveLetter% %Vol1%:\ /mir /xf autounattend.xml VOL1_M_MEDIA.WIM /xd sources "; Chr$(34); "system volume information"; Chr$(34); " $recycle.bin /256 /r:0 > NUL"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler1"
    Print #ff, "robocopy %MountedImageDriveLetter%\sources %Vol1%:\sources boot.wim /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler1"
    Print #ff, "robocopy %MountedImageDriveLetter%\sources %Vol2%:\sources /mir /xf boot.wim /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy %MountedImageDriveLetter%\boot %Vol2%:\boot /mir /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy %MountedImageDriveLetter%\efi %Vol2%:\efi /mir /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy %MountedImageDriveLetter%\support %Vol2%:\support /mir /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, ""
    Print #ff, "if not exist %Vol2%:\sources\ei.cfg ("
    Print #ff, "echo [CHANNEL] > %Vol2%:\sources\ei.cfg"
    Print #ff, "echo Retail >> %Vol2%:\sources\ei.cfg"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "goto DoneCopying"
    Print #ff, ""
    Print #ff, ":CopyDual"
    Print #ff, ""
    Print #ff, "if %SourceAF% == Y ("
    Print #ff, "robocopy %MountedImageDriveLetter% %Vol1%:\ /mir /xf VOL1_M_MEDIA.WIM /xd sources x64 x86 "; Chr$(34); "system volume information"; Chr$(34); " $recycle.bin /256 /r:0 > NUL"
    Print #ff, ") ELSE ("
    Print #ff, "robocopy %MountedImageDriveLetter% %Vol1%:\ /mir /xf autounattend.xml VOL1_M_MEDIA.WIM /xd sources x64 x86 "; Chr$(34); "system volume information"; Chr$(34); " $recycle.bin /256 /r:0 > NUL"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler1"
    Print #ff, "robocopy %MountedImageDriveLetter%\x64\sources %Vol1%:\x64\sources boot.wim /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler1"
    Print #ff, "robocopy %MountedImageDriveLetter%\x86\sources %Vol1%:\x86\sources boot.wim /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler1"
    Print #ff, "robocopy "; Chr$(34); "%MountedImageDriveLetter%\x64\sources "; Chr$(34); " %Vol2%:\x64\sources /mir /xf boot.wim /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy "; Chr$(34); "%MountedImageDriveLetter%\x86\sources "; Chr$(34); " %Vol2%:\x86\sources /mir /xf boot.wim /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy "; Chr$(34); "%MountedImageDriveLetter%\x64\boot "; Chr$(34); " %Vol2%:\x64\boot /mir /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy "; Chr$(34); "%MountedImageDriveLetter%\x86\boot "; Chr$(34); " %Vol2%:\x86\boot /mir /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy "; Chr$(34); "%MountedImageDriveLetter%\x64\efi "; Chr$(34); " %Vol2%:\x64\efi /mir /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy "; Chr$(34); "%MountedImageDriveLetter%\x86\efi "; Chr$(34); " %Vol2%:\x86\efi /mir /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy "; Chr$(34); "%MountedImageDriveLetter%\x64\support "; Chr$(34); " %Vol2%:\x64\support /mir /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, "robocopy "; Chr$(34); "%MountedImageDriveLetter%\x86\support "; Chr$(34); " %Vol2%:\x86\support /mir /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler2"
    Print #ff, ""
    Print #ff, "if not exist %Vol2%:\x64\sources\ei.cfg ("
    Print #ff, "echo [CHANNEL] > %Vol2%:\x64\sources\ei.cfg"
    Print #ff, "echo Retail >> %Vol2%:\x64\sources\ei.cfg"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if not exist %Vol2%:\x86\sources\ei.cfg ("
    Print #ff, "echo [CHANNEL] > %Vol2%:\x86\sources\ei.cfg"
    Print #ff, "echo Retail >> %Vol2%:\x86\sources\ei.cfg"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "goto DoneCopying"
    Print #ff, ""
    Print #ff, ":CopyPE"
    Print #ff, ""
    Print #ff, "robocopy %MountedImageDriveLetter% %Vol1%:\ /mir /xf VOL1_M_MEDIA.WIM /xd "; Chr$(34); "system volume information"; Chr$(34); " $recycle.bin /256 /r:0 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto ErrorHandler1"
    Print #ff, "goto DoneCopying"
    Print #ff, ""
    Print #ff, ":DoneCopying"
    Print #ff, ""
    Print #ff, "REM All the files have been copied except for the answer file. If the user selected an answer file"
    Print #ff, "REM copy it now."
    Print #ff, ""
    Print #ff, "if %UserAF% == Y ("
    Print #ff, ""
    Print #ff, "copy /Y "; Chr$(34); "%af_UserAnswerFile%"; Chr$(34); " %Vol1%:\autounattend.xml > NUL"
    Print #ff, "echo %af_FileNameOnly% > %Vol1%:\Answer_File.txt"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "if %SourceAF% == Y ("
    Print #ff, "echo Answer file from original source media has been copied >  %Vol1%:\Answer_File.txt"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: Copying of files has been completed ::"
    Print #ff, "echo ::                                     ::"
    Print #ff, "echo ::      Dismounting the image          ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'%SourceISOImage%'"; Chr$(34); ""; Chr$(34); " > NUL"
    Print #ff, ""
    Print #ff, "if %ImageType%==Single goto SummarizeWinImage"
    Print #ff, "if %ImageType%==Dual goto SummarizeWinImage"
    Print #ff, "if %ImageType%==PE goto SummarizePEImage"
    Print #ff, ""
    Print #ff, ":SummarizeWinImage"
    Print #ff, ""

    ' Commenting out the section below that patches for Secure Boot. It is no longer needed with the ADK and Windows
    ' PE release of May, 2024

    'Print #ff, "cls"
    'Print #ff, "REM Start of routine to patch for Secure Boot mitigations"
    'Print #ff, ""
    'Print #ff, "REM If an X:\Windows folder it exists, this indicates that the program was run by booting from this disk and not"
    'Print #ff, "REM run from within Windows. In this case, display the message below and skip trying to patch for Secure Boot."
    'Print #ff, ""
    'Print #ff, "If exist X:\Windows ("
    'Print #ff, "echo."
    'Print #ff, "echo Because you booted from a bootable disk and not into Windows, we cannot perform a check to see if"
    'Print #ff, "echo your system has Secure Boot mitigations for the BlackLotus UEFI Bootkit vulnerability. It is possible"
    'Print #ff, "echo that this disk may not boot from a system with those mitigations installed. If possible, it is suggested"
    'Print #ff, "echo that you boot into Windows on a system with those mitigations applied, revert the disk back to the default"
    'Print #ff, "echo state, and then reselect the item to boot. If you plan to boot this disk on a system without those"
    'Print #ff, "echo mitigations, then it shouod boot just fine."
    'Print #ff, "goto DonePatchingB"
    'Print #ff, ")"
    'Print #ff, ""
    'Print #ff, "echo *************************************************************************************"
    'Print #ff, "echo *                              Applying Secure Boot Fixes                           *"
    'Print #ff, "echo *                                                                                   *"
    'Print #ff, "echo * Checking to see if this system has BlackLotus UEFI Bootkit mitigations installed. *"
    'Print #ff, "echo * If it does, we will use this to patch this media to work on systems with those    *"
    'Print #ff, "echo * mitigations applied. Without this update, this media may not boot on systems with *"
    'Print #ff, "echo * those mitigations. If this disk will not boot on some systems, please do this:    *"
    'Print #ff, "echo *                                                                                   *"
    'Print #ff, "echo * 1) Rerun this program on a system with those mitigations installed and perform a  *"
    'Print #ff, "echo *    REFRESH operation to refresh the boot information on this disk.                *"
    'Print #ff, "echo *                                                                                   *"
    'Print #ff, "echo * 2) If this program asks if you want to use a previously created BOOT.WIM file,    *"
    'Print #ff, "echo *    choose NOT to use a previously created file.                                   *"
    'Print #ff, "echo *************************************************************************************"
    'Print #ff, "echo."
    'Print #ff, "REM Check to make sure that the first of two mitigations are applied to this system."
    'Print #ff, "REM The first mitigation adds the "; Chr$(34); "Windows UEFI CA 2023"; Chr$(34); " certificate to the UEFI "; Chr$(34); "Secure Boot Signature Database"; Chr$(34); " (DB)."
    'Print #ff, "REM By adding this certificate to the DB, the device firmware will trust boot applications signed by this certificate."
    'Print #ff, ""
    'Print #ff, "for /f %%a in ('powershell "; Chr$(34); "[System.Text.Encoding]::ASCII.GetString((Get-SecureBootUEFI db).bytes) -match 'Windows UEFI CA 2023'"; Chr$(34); "') do set "; Chr$(34); "PowerShellOutput=%%a"; Chr$(34); ""
    'Print #ff, ""
    'Print #ff, "if "; Chr$(34); "%PowerShellOutput%"; Chr$(34); "=="; Chr$(34); "True"; Chr$(34); " ("
    'Print #ff, "goto CheckCondition2B"
    'Print #ff, ") else ("
    'Print #ff, "goto NotInstalledB"
    'Print #ff, ")"
    'Print #ff, ""
    'Print #ff, ":CheckCondition2B"
    'Print #ff, ""
    'Print #ff, "REM Check to make sure that the second of two mitigations are applied to this system."
    'Print #ff, "REM The UEFI Forbidden List (DBX) is used to block untrusted UEFI modules from loading. The second mitigation updates"
    'Print #ff, "REM the DBX by adding the "; Chr$(34); "Windows Production CA 2011"; Chr$(34); " certificate to the DBX. This will cause all boot managers signed by"
    'Print #ff, "REM this certificate to no longer be trusted."
    'Print #ff, ""
    'Print #ff, "for /f %%a in ('powershell "; Chr$(34); "[System.Text.Encoding]::ASCII.GetString((Get-SecureBootUEFI dbx).bytes) -match 'Microsoft Windows Production PCA 2011'"; Chr$(34); "') do set "; Chr$(34); "PowerShellOutput=%%a"; Chr$(34); ""
    'Print #ff, ""
    'Print #ff, "if "; Chr$(34); "%PowerShellOutput%"; Chr$(34); "=="; Chr$(34); "True"; Chr$(34); " ("
    'Print #ff, "goto Condition2TrueB"
    'Print #ff, ") else ("
    'Print #ff, "goto NotInstalledB"
    'Print #ff, ")"
    'Print #ff, ""
    'Print #ff, ":Condition2TrueB"
    'Print #ff, ""
    'Print #ff, "REM We have verified that the mitigations for the BlackLotus UEFI Bootkit are installed on this system. We will now update"
    'Print #ff, "REM the boot media to ensure that it can be successfully booted on this system."
    'Print #ff, ""
    'Print #ff, "REM Make sure that the files on the destination disk are not read only"
    'Print #ff, ""
    'Print #ff, "attrib -r "; Letter$(1); ":\*.* /s /d > NUL 2>&1"
    'Print #ff, "COPY /Y /B "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD.BAK > NUL 2>&1"
    'Print #ff, "bcdboot c:\windows /f UEFI /s "; Letter$(1); ": /bootex > NUL"
    'Print #ff, "COPY /Y /B "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD.BAK "; Letter$(1); ":\EFI\MICROSOFT\BOOT\BCD > NUL"
    'Print #ff, "echo This system DOES have the mitigations installed. Media has been patched."
    'Print #ff, "goto DonePatchingB"
    'Print #ff, ""
    'Print #ff, ":NotInstalledB"
    'Print #ff, ""
    'Print #ff, "REM We arrive here if the mitigations are not installed on this system or when updates are doing being installed."
    'Print #ff, "echo This system DOES NOT have the mitigations installed. Media has NOT been patched."
    'Print #ff, ""
    'Print #ff, ":DonePatchingB"
    'Print #ff, ""
    'Print #ff, "REM Done with routine to patch for Secure Boot mitigations"
    'Print #ff, ""
    Print #ff, "echo."
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: Creation of the disk has completed. You can now boot the image you selected from this disk. ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "if exist %Vol1%:\autounattend.xml ("
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: CAUTION^^! Your disk includes an unattended answer file ^(autounattend.xml^). If your ::"
    Print #ff, "echo :: system is configured to boot from the disk, Windows installation will begin       ::"
    Print #ff, "echo :: automatically. If your answer file is configured to wipe a disk^(s^), then this     ::"
    Print #ff, "echo :: will happen automatically with no warning.                                        ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, "if %SourceAF% == Y ("
    Print #ff, "echo NOTE: The answer file being used was copied from the original source image."
    Print #ff, "echo."
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "goto End_Config"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "REM Normally, Vol2 should not have an answer file, but we check just in case"
    Print #ff, "REM the user added one manually."
    Print #ff, ""
    Print #ff, "if exist %Vol2%:\autounattend.xml ("
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: CAUTION^^! Your disk includes an unattended answer file ^(autounattend.xml^). If your ::"
    Print #ff, "echo :: system is configured to boot from the disk, Windows installation will begin       ::"
    Print #ff, "echo :: automatically. If your answer file is configured to wipe a disk^(s^), then this     ::"
    Print #ff, "echo :: will happen automatically with no warning.                                        ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, "goto End_Config"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "goto End_Config"
    Print #ff, ""
    Print #ff, ":SummarizePEImage"
    Print #ff, ""
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: Creation of the disk has completed. You can now boot the image you selected from this disk. ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, "goto End_Config"
    Print #ff, ""
    Print #ff, ":::::::::::::::::::::::::::::"
    Print #ff, ":: Error Handling Routines ::"
    Print #ff, ":::::::::::::::::::::::::::::"
    Print #ff, ""
    Print #ff, ":ErrorHandler1"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo There was an error copying files to volume #1. Please verify that volume #1 has sufficient space available."
    Print #ff, "echo Please correct this situation."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto End"
    Print #ff, ""
    Print #ff, ":ErrorHandler2"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo There was an error copying files to volume #2. Please verify that volume #2 has sufficient space available."
    Print #ff, "echo Please correct this situation."
    Print #ff, "echo."
    Print #ff, "goto End"
    Print #ff, ""
    Print #ff, ":End_Config"
    Print #ff, ""
    Print #ff, "echo The program will now end."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto End"
    Close #ff
Next x

For x = 1 To 2

    Select Case x
        Case 1
            ff = FreeFile
            Open "startnet.cmd" For Append As #ff
        Case 2
            ff = FreeFile
            Open "Config_UFD.bat" For Append As #ff
    End Select

    Print #ff, ""
    Print #ff, ":Restore"
    Print #ff, ""
    Print #ff, "REM Need to get the drive letters for the 2 volumes of the boot media"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo Retrieving drive letters. Please standby..."
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "REM Finding drive letter for Volume 1"
    Print #ff, ""
    Print #ff, "FOR %%a IN (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do ("
    Print #ff, ""
    Print #ff, "IF exist %%a:\VOL1_M_MEDIA.WIM ("
    Print #ff, "set vol1=%%a"
    Print #ff, "goto Vol1Found2"
    Print #ff, "   )"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "echo We could not find Volume 1. The program will now end."
    Print #ff, "pause"
    Print #ff, "goto End"
    Print #ff, ""
    Print #ff, ":Vol1Found2"
    Print #ff, ""
    Print #ff, "REM Finding drive letter for Volume 2"
    Print #ff, ""
    Print #ff, "FOR %%a IN (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do ("
    Print #ff, ""
    Print #ff, "IF exist %%a:\VOL2_M_MEDIA.WIM ("
    Print #ff, "set vol2=%%a"
    Print #ff, "goto Vol2Found2"
    Print #ff, "   )"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "echo We could not find Volume 2. The program will now end."
    Print #ff, "pause"
    Print #ff, "goto End"
    Print #ff, ""
    Print #ff, ":Vol2Found2"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo :::::::::::::::::::::::::"
    Print #ff, "echo :: Restoring volume #1 ::"
    Print #ff, "echo :::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "robocopy %Vol2%:\PE_Backup %Vol1%:\ /mir /xf VOL1_M_MEDIA.WIM /xd "; Chr$(34); "system volume information"; Chr$(34); " $recycle.bin > NUL"
    Print #ff, ""
    Print #ff, "REM Checking for the existance of an answer file so that we can warn the user later"
    Print #ff, ""
    Print #ff, "if exist %Vol1%:\autounattend.xml ("
    Print #ff, "set AutoVol1Exist=Y"
    Print #ff, ") ELSE ("
    Print #ff, "set AutoVol1Exist=N"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "echo :::::::::::::::::::::::::"
    Print #ff, "echo :: Restoring volume #2 ::"
    Print #ff, "echo :::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "RD /Q /S %Vol2%:\Sources > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\x64\Sources > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\x86\Sources > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\boot > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\x64\boot > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\x86\boot > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\efi > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\x64\efi > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\x86\efi > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\support > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\x64\support > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\x86\support > NUL 2>&1"
    Print #ff, "RD /Q /S %Vol2%:\PE_Backup > NUL 2>&1"
    Print #ff, ""
    Print #ff, "REM Checking for the existance of an answer file so that we can warn the user later"
    Print #ff, ""
    Print #ff, "if exist %Vol2%:\autounattend.xml ("
    Print #ff, "set AutoVol2Exist=Y"
    Print #ff, ") ELSE ("
    Print #ff, "set AutoVol2Exist=N"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "echo ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: Recreating tags to identify the disk as a multi image disk ::"
    Print #ff, "echo ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, "echo. > %Vol1%:\VOL1_M_MEDIA.WIM"
    Print #ff, "echo. > %Vol2%:\VOL2_M_MEDIA.WIM"
    Print #ff, "del %Vol2%:\Boot_Image.txt > NUL"
    Print #ff, "echo ::::::::::::::::::::::::::"
    Print #ff, "echo :: Restoration complete ::"
    Print #ff, "echo ::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, ""
    Print #ff, "REM Checking the first volume to see if an autounattend.xml answer file is present."
    Print #ff, ""
    Print #ff, ":HandleVol1AnswerFile"
    Print #ff, ""
    Print #ff, "if %AutoVol1Exist%==N goto HandleVol2AnswerFile"
    Print #ff, "cls"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: CAUTION^^! Your disk includes an unattended answer file ^(autounattend.xml^). If your ::"
    Print #ff, "echo :: system is configured to boot from the disk, Windows installation will begin       ::"
    Print #ff, "echo :: automatically. If your answer file is configured to wipe a disk^(s^), then this     ::"
    Print #ff, "echo :: will happen automatically with no warning.                                        ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, "set /P DelFile="; Chr$(34); "Do you want to delete this file? "; Chr$(34)
    Print #ff, "if [%DelFile%]==[] goto HandleVol1AnswerFile"
    Print #ff, "set DelFile=%DelFile:~0,1%"
    Print #ff, "call ::TOUPPERCASE DelFile"
    Print #ff, "if %DelFile%==Y goto DelFromVol1"
    Print #ff, "if %DelFile% NEQ N goto HandleVol1AnswerFile"
    Print #ff, "echo."
    Print #ff, "echo The answer file will not be deleted."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto :HandleVol2AnswerFile"
    Print #ff, ""
    Print #ff, ":DelFromVol1"
    Print #ff, ""
    Print #ff, "attrib -h -s -r %Vol1%:\autounattend.xml > NUL"
    Print #ff, "del %Vol1%:\autounattend.xml > NUL"
    Print #ff, "echo."
    Print #ff, "echo The answer file has been deleted."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, ""
    Print #ff, "REM Checking the second volume to see if an autounattend.xml answer file is present."
    Print #ff, ""
    Print #ff, ":HandleVol2AnswerFile"
    Print #ff, ""
    Print #ff, "if %AutoVol2Exist%==N goto HandleAnswerFilesDone"
    Print #ff, "cls"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo :: CAUTION^^! Your disk includes an unattended answer file ^(autounattend.xml^). If your ::"
    Print #ff, "echo :: system is configured to boot from the disk, Windows installation will begin       ::"
    Print #ff, "echo :: automatically. If your answer file is configured to wipe a disk^(s^), then this     ::"
    Print #ff, "echo :: will happen automatically with no warning.                                        ::"
    Print #ff, "echo :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, "echo."
    Print #ff, "set /P DelFile="; Chr$(34); "Do you want to delete this file? "; Chr$(34)
    Print #ff, "if [%DelFile%]==[] goto HandleVol2AnswerFile"
    Print #ff, "set DelFile=%DelFile:~0,1%"
    Print #ff, "call ::TOUPPERCASE DelFile"
    Print #ff, "if %DelFile%==Y goto DelFromVol2"
    Print #ff, "if %DelFile% NEQ N goto HandleVol2AnswerFile"
    Print #ff, "echo."
    Print #ff, "echo The answer file will not be deleted."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto :HandleAnswerFilesDone"
    Print #ff, ""
    Print #ff, ":DelFromVol2"
    Print #ff, ""
    Print #ff, "attrib -h -s -r %Vol2%:\autounattend.xml > NUL"
    Print #ff, "del %Vol2%:\autounattend.xml > NUL"
    Print #ff, "echo."
    Print #ff, "echo The answer file has been deleted."
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, ""
    Print #ff, ":HandleAnswerFilesDone"
    Print #ff, ""
    Print #ff, "cls"
    Print #ff, "echo The disk has been restored to the original state and is ready to be reconfigured by running the Config_UFD.bat file"
    Print #ff, "echo located on the second volume."
    Print #ff, ""
    Print #ff, "REM If an X:\Windows folder it exists, this indicates that the program was run by booting from this disk and"
    Print #ff, "REM not run from within Windows. In this case, display the message below."
    Print #ff, ""
    Print #ff, "If exist X:\Windows ("
    Print #ff, "echo."
    Print #ff, "echo Before you run this batch file, you should first reboot the system to allow the current changes to take effect."
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "echo."
    Print #ff, "pause"
    Print #ff, "goto End"
    Print #ff, ""
    Print #ff, ":::::::::::::::"
    Print #ff, ":: FUNCTIONS ::"
    Print #ff, ":::::::::::::::"
    Print #ff, ""
    Print #ff, ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, ":: The following is a function that can be called to convert ::"
    Print #ff, ":: the contents of a variable to uppercase characters.       ::"
    Print #ff, "::                                                           ::"
    Print #ff, ":: To use this function, call this function and pass it the  ::"
    Print #ff, ":: name of the variable to convert. In the example below we  ::"
    Print #ff, ":: are passing str as the variable name.                     ::"
    Print #ff, "::                                                           ::"
    Print #ff, ":: Example: call ::TOUPPERCASE str                           ::"
    Print #ff, ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #ff, ""
    Print #ff, ":TOUPPERCASE"
    Print #ff, "if not defined %~1 exit /b"
    Print #ff, "for %%a in ("; Chr$(34); "a=A"; Chr$(34); " "; Chr$(34); "b=B"; Chr$(34); " "; Chr$(34); "c=C"; Chr$(34); " "; Chr$(34); "d=D"; Chr$(34); " "; Chr$(34); "e=E"; Chr$(34); " "; Chr$(34); "f=F"; Chr$(34); " "; Chr$(34); "g=G"; Chr$(34); " "; Chr$(34); "h=H"; Chr$(34); " "; Chr$(34); "i=I"; Chr$(34); " "; Chr$(34); "j=J"; Chr$(34); " "; Chr$(34); "k=K"; Chr$(34); " "; Chr$(34); "l=L"; Chr$(34); " "; Chr$(34); "m=M"; Chr$(34); " "; Chr$(34); "n=N"; Chr$(34); " "; Chr$(34); "o=O"; Chr$(34); " "; Chr$(34); "p=P"; Chr$(34); " "; Chr$(34); "q=Q"; Chr$(34); " "; Chr$(34); "r=R"; Chr$(34); " "; Chr$(34); "s=S"; Chr$(34); " "; Chr$(34); "t=T"; Chr$(34); " "; Chr$(34); "u=U"; Chr$(34); " "; Chr$(34); "v=V"; Chr$(34); " "; Chr$(34); "w=W"; Chr$(34); " "; Chr$(34); "x=X"; Chr$(34); " "; Chr$(34); "y=Y"; Chr$(34); " "; Chr$(34); "z=Z"; Chr$(34); " "; Chr$(34); "ä=Ä"; Chr$(34); " "; Chr$(34); "ö=Ö"; Chr$(34); " "; Chr$(34); "ü=Ü"; Chr$(34); ") do ("
    Print #ff, "call set %~1=%%%~1:%%~a%%"
    Print #ff, ")"
    Print #ff, ""
    Print #ff, "goto :eof"
    Print #ff, ""
    Print #ff, "::::::::::::::::::::::"
    Print #ff, ":: END OF FUNCTIONS ::"
    Print #ff, "::::::::::::::::::::::"
    Print #ff, ""
    Print #ff, ":End"
    Close #ff
Next x

' Run the batch file that will create the boot disk for us

Shell "create_disk.bat"

' Cleanup the files created earlier. They are no longer needed.

If _FileExists("create_disk.bat") Then
    Kill "create_disk.bat"
End If

If _FileExists("startnet.cmd") Then
    Kill "startnet.cmd"
End If

If _FileExists("config_ufd.bat") Then
    Kill "config_ufd.bat"
End If

' Display usage instructions

' We want to make certain that the usage instructions are displayed to the user. As a result, we will clear the keyboard
' buffer to avoid having any accidentally pressed keys from skipping past the below messages.
'
' Since we may arrive at this screen too quickly while a key is still depressed, we add a brief delay here.

_Delay 2
_KeyClear

' Now proceed with display of the messages.

Cls
Print "Disk creation has been completed."
Print
Color 0, 10: Print "Usage Instructions": Color 15
Print
Print "On the 2nd partition of the disk that we just created (the one with the volume label ";: Color 0, 14: Print VolLabel$(2);: Color 15: Print "),"
Print "you will find a folder called ";: Color 0, 14: Print "ISO Images";: Color 15: Print ". Place any Windows ISO images that you may want to use in this folder. In"
Print "addition, you can place images of Windows PE or Windows RE based media here. For example, a Macrium Reflect boot disk"
Print "or a Windows recovery disk."
Print
Print "You will also find a folder called ";: Color 0, 14: Print "Other";: Color 15: Print " here. This is a good place to save any other files, for example, a batch file"
Print "to bypass Windows 11 system requirements, notes you may need, etc."
Print
Print "Place any unattended answer files that you may wish to use in the folder named ";: Color 0, 14: Print "Answer Files";: Color 15: Print ". You can call these"
Print "anything you wish but you should end the name with the extension ";: Color 0, 14: Print ".xml";: Color 15: Print ". When you select an answer file for use, it will"
Print "be copied to the destination as ";: Color 0, 14: Print "autounattend.xml";: Color 15: Print "."
Print
Print "Start by running ";: Color 0, 14: Print "Config_UFD.bat";: Color 15: Print " from ";: Color 0, 14: Print VolLabel$(2);: Color 15: Print "."
Print
Print "This will configure your disk to boot the ISO image and optional answer file of your choosing. Note that you can also"
Print "boot from the disk and then select the ISO image to boot. This helps if you cannot boot into Windows."
Print
Print "Please select an answer file only for a Windows image, not for any Windows PE or Windows RE based images."
Print
Color 0, 10: Print "Booting from the Drive": Color 15
Print
Print "If you have not run ";: Color 0, 14: Print "Config_UFD.bat";: Color 15: Print " to select an image to boot, on the first boot, you will be shown a list of all the"
Print "ISO images that you placed in the ISO Images folder. You will then select the image to be made bootable. It will then"
Print "reconfigure the disk to boot the selected image. You are now ready to boot the selected image. If you ran the batch"
Print "file from within Windows and selected a image to boot, then your disk is already configured to boot that image."
Pause
Cls
Color 0, 10: Print "Cleanup": Color 15
Print
Print "When you are done, run the ";: Color 0, 14: Print "Config_UFD.bat";: Color 15: Print " located on the second partition. This will restore the disk to the"
Print "original state, ready for you to select a new image to boot. Please note that you can run this batch file from within"
Print "Windows on another system, or, if you do not have access to a system that can boot Windows, you can run this batch file"
Print "from a command line while booted from this disk. You can press ";: Color 0, 14: Print "SHIFT + F10";: Color 15: Print " to open a command prompt from which you can"
Print "run the batch file."
Print
Print "TIP: If you are trying to locate the batch file while booted from the disk, it can be a little awkward because File"
Print "Explorer is not available. As a workaround, from the command prompt, type NOTEPAD and press ENTER. Change Text Documents"
Print "(*.txt) to All files (*.*), and then select File > Open to view all the disks on the system. Look for the volume with"
Print "the volume label ";: Color 0, 14: Print VolLabel$(2);: Color 15: Print "."
Print
Print "Once you have located the correct drive, double-click it to open it, then right-click on the batch file called"
Color 0, 14: Print "Config_UFD.bat";: Color 15: Print " (the .bat may not show in the file name), and select the option to run as administrator."
Print
Color 0, 10: Print "A Note About Answer Files": Color 15
Print
Print "Leaving unattended answer files on any volume of a bootable disk can potentially be dangerous. If you accidentally boot"
Print "from a disk with an answer file, there is the potential that one or more disks may be wiped clean and Windows will"
Print "begin installation automatically. What exactly will happen depends upon the configuration of the answer file but this"
Print "is something to be cautious about. In an effort to prevent accidents, the program will notify you about the presence"
Print "of an answer file on either of the first two partitions and will offer to remove them. Note that when you revert the"
Print "disk back to the original state, if there was no answer file on the original first volume, any answer file that was"
Print "added will be automatically removed."
Pause

ChDir ProgramStartDir$: GoTo BeginProgram

' Local subroutines follow

' Subroutine - Shows patition information.

ShowPartitionSizes:

Cls
Color 0, 14
Print "*******************"
Print "* PARTITION SIZES *"
Print "*******************"
Color 15
Print
Print "Partition #1:";

If PartitionSize$(1) = "" Then
    Color 14, 4: Print " NOT YET DEFINED ";: Color 15
Else
    DisplayUnit$ = "MB"
    TempValue = Val(PartitionSize$(1))
    If TempValue >= 1024 Then
        TempValue = TempValue / 1024
        DisplayUnit$ = "GB"
    End If
    If TempValue >= 1024 Then
        TempValue = TempValue / 1024
        DisplayUnit$ = "TB"
    End If
    Color 0, 10: Print Using "####.##"; TempValue;: Print " "; DisplayUnit$; " ";: Color 15
End If

Print " - Holds boot files. Assign 2.5 GB if you have enough space."
Print "Partition #2:";

If PartitionSize$(2) = "" Then

    Select Case TotalPartitions
        Case 2
            Color 0, 10: Print " All space not assigned to the first partition ";: Color 15
        Case 3, 4
            Color 14, 4: Print " NOT YET DEFINED ";: Color 15
    End Select

Else
    DisplayUnit$ = "MB"
    TempValue = Val(PartitionSize$(2))

    If TempValue >= 1024 Then
        TempValue = TempValue / 1024
        DisplayUnit$ = "GB"
    End If

    If TempValue >= 1024 Then
        TempValue = TempValue / 1024
        DisplayUnit$ = "TB"
    End If

    Color 0, 10: Print Using "####.##"; TempValue;: Print " "; DisplayUnit$; " ";: Color 15
End If

Print " - Make large enough to hold your image file(s)"

Select Case TotalPartitions
    Case 4
        Print "Partition #3:";
        If PartitionSize$(3) = "" Then
            Color 14, 4: Print " NOT YET DEFINED ": Color 15
        Else

            DisplayUnit$ = "MB"
            TempValue = Val(PartitionSize$(3))
            If TempValue >= 1024 Then
                TempValue = TempValue / 1024
                DisplayUnit$ = "GB"
            End If
            If TempValue >= 1024 Then
                TempValue = TempValue / 1024
                DisplayUnit$ = "TB"
            End If
            Color 0, 10: Print Using "####.##"; TempValue;: Print " "; DisplayUnit$; " ": Color 15
        End If
        Print "Partition #4:";: Color 0, 10: Print " All space not assigned to the first three partitions ": Color 15
    Case 3
        Print "Partition #3:";: Color 0, 10: Print " All space not assigned to the first two partitions ": Color 15
End Select

Return

' Local subroutine
' Display a list of disks and ask which one to use. If user needs more detail on a disk, display that detail.

SelectDisk:

Cls
Print "Building a list of disks in the system..."

GetDiskDetails

' NOTE: Since a Disk ID of 0 may be valid on some systems (Disk 0 is not always the boot disk), we cannot assume that
' a value of 0 is invalid. If the user simply hits ENTER in response to a numerical entry, that value would be 0,
' but we don't want to consider an ENTER without an actual response to be valid. For this reason, we ask for a string
' input where we can check for a nul input and flag that as invalid. We can then convert the string response into a
' number.

' At the end of this routine DiskID will hold the ID of the disk chosen by the user.

' Init varaibles

Temp$ = ""
ValidDisk = 0

AskForDiskID:

Do
    Cls
    Print "Please note the number of the disk that you want to make bootable."
    Print
    Color 0, 14: Print ListOfDisks$: Color 15
    Print "Description for each disk:"
    Print
    Color 0, 14

    For x = 1 To NumberOfDisks
        Print "Disk"; DiskIDList(x); "Description: "; DiskDetail$(x)
    Next x

    Color 15
    Print
    Input "Enter the disk number of the disk you want to make bootable: ", Temp$
Loop While Temp$ = ""

DiskID = Val(Temp$)

' Since typing a non-numeric response would yield a val of 0, we need to do a check to see if what the user
' entered was really a 0 or something else.

If DiskID = 0 And Temp$ <> "0" Then
    GoTo AskForDiskID
End If

For x = 1 To NumberOfDisks

    If DiskID = DiskIDList(x) Then
        ValidDisk = 1
        GoTo ValidDiskCheckDone
    End If

Next x

ValidDiskCheckDone:

If ValidDisk = 0 Then
    Cls
    Print
    Color 14, 4: Print "Invalid Disk Selected!";: Color 15: Print " Please enter one of the disk numbers shown by the program to be valid."
    Pause
    GoTo AskForDiskID
End If

Print
Print "You have selected the following disk: ";: Color 0, 14: Print "Disk"; DiskID; "- "; DiskDetail$(x);: Color 15
Print
Input "Is this correct"; Temp$
YesOrNo Temp$
Temp$ = YN$

Select Case Temp$
    Case "X", "N"
        GoTo AskForDiskID
    Case "Y"
        Exit Select
End Select

Return

Generic_Partition_Size:

' Ask for the size to make a partition. Accepts input in the along with M for Megabytes, G for Gigabytes, and T for Terabytes.
' After running this routine, the variable "ParSizeInMB$" will hold the size of the partition in MB as a string. This is useful
' where we want to print references to the size without leading or training spaces.
'
' Note that text should be displayed to user before coming to this routine to explain what partition a size is being sought. This
' routine only displays a generic prompt for the size of the partition so that it can be used universally.

RedoPartitionSize_2:

ParSizeInMB$ = "" ' Set initial value

Print "Enter the size below followed by "; Chr$(34); "M"; Chr$(34); " for Megabytes, "; Chr$(34); "G"; Chr$(34); " for Gigabytes, or "; Chr$(34); "T"; Chr$(34); " for Terabytes."
Print
Print "Examples: 500M, 20G, 1T, 700m, 1g"
Print
Print "Enter the size for this partition: ";
Input "", TempPartitionSize$
TempUnit$ = UCase$(Right$(TempPartitionSize$, 1))
TempValue = Val(TempPartitionSize$)
Select Case TempUnit$
    Case "M"
        ParSizeInMB$ = Str$(TempValue)
        GoTo PartitionUnitsValid_2
    Case "G"
        ParSizeInMB$ = Str$(TempValue * 1024)
        GoTo PartitionUnitsValid_2
    Case "T"
        ParSizeInMB$ = Str$(TempValue * 1048576)
        GoTo PartitionUnitsValid_2
    Case Else

        ' A valid entry was not made. We will return to the calling section of code. It's up to that code to reprompt for
        'valid information and then call this routine again.

        Return

End Select

PartitionUnitsValid_2:

If (Val(ParSizeInMB$)) <= 100 Then
    Cls
    Color 14, 4: Print "This program expects a minimum partition size of 100 MB.": Color 15
    Pause

    ' A valid entry was not made. We will return to the calling section of code. It's up to that code to reprompt for
    'valid information and then call this routine again.

    Return

End If

Return

' Get drive letter to assign to each partition
' The user can choose to manually assign drive letters to the partitions being created on the bootable media
' or allow the program to automatically assign drive letters.

SelectAutoOrManual:

ReDim Letter(TotalPartitions) As String

Cls
Print "The program will automatically assign drive letters to the"; TotalPartitions; "partitions that will be created on the drive. However,"
Print "if you prefer, you can manually assign drive letters."
Print
Input "Do you want to manually assign drive letters"; ManualAssignment$

YesOrNo ManualAssignment$

Select Case YN$
    Case "Y"
        GoTo ManualAssign
    Case "N"
        GoTo AutoAssign
    Case "X"
        GoTo SelectAutoOrManual
End Select

ManualAssign:

' Allow the user to manually choose drive letters

For x = 1 To TotalPartitions
    Do

        RepeatLetter:

        Cls
        Print "For all"; TotalPartitions; "partitions, enter the drive letter to assign. Enter only the letter without a colon (:)."
        Print
        Letter$(x) = "" ' Set initial value
        Print "Enter the drive letter for partition #";: Color 0, 10: Print x;: Color 15: Print ": ";
        Input "", Letter$(x)
    Loop While Letter$(x) = ""

    Letter$(x) = UCase$(Letter$(x))

    If (Len(Letter$(x)) > 1) Or (Letter$(x)) = "" Or ((Asc(Letter$(x))) < 65) Or ((Asc(Letter$(x))) > 90) Then
        Print
        Color 14, 4: Print "That was not a valid entry. Please try again.": Color 15
        Print
        GoTo RepeatLetter
    End If

    If _DirExists(Letter$(x) + ":") Then
        Print
        Color 14, 4: Print "That drive letter is already in use.": Color 15
        Pause
        GoTo RepeatLetter
    End If

Next x

GoTo LetterAssignmentDone

AutoAssign:

' Automatically assign drive letters

' To auto assign drive letters, we go through a loop checking to see if drive letters C:, D:, E:, etc. are already in use.
' If in use, we move on to the next drive letter. After the first free drive letter is picked, we repeat the same process
' but we resume checking for free drive letters with the letter after the last one we just assigned.

' Before we search for available drive letters, we will call a subroutine that looks for removable media with volumes that have
' a status of "Unusable". This happens when a removable media drive such as a USB Flash Drive (UFD) has had a "clean" operation
' performed on it in Diskpart. The drive now has now partitions but it still shows up in File Explorer with a drive letter. This
' can cause difficulties for drive letter detection.

Print
Print "Looking for available drive letters."

CleanVols

LettersAssigned = 0 ' Keep track of how many drive letters were assigned. Once equal to the number of partitions, we are done.

Restore DriveLetterData

For y = 1 To 24
    Read Letter$(LettersAssigned + 1)
    Cmd$ = "dir " + Letter$(LettersAssigned + 1) + ":\ > DriveStatus.txt 2>&1"
    Shell Cmd$
    ff = FreeFile
    Open "DriveStatus.txt" For Input As #ff
    FileLength = LOF(ff)
    Temp$ = Input$(FileLength, ff)
    Close #ff
    Kill "DriveStatus.txt"

    If InStr(Temp$, "The system cannot find the path specified") Then
        LettersAssigned = LettersAssigned + 1
    End If

    If LettersAssigned = TotalPartitions Then GoTo LetterAssignmentDone

Next y

' The FOR loop should only complete if we run out of drive letters. We need to warn the user about this and how they can correct
' the issue. The program will then end.

Cls
Print "Not enough drive letters were available to assign to all partititions!"
Print
Print "Solution: Please free up some drive letters and then re-run this program."
Print "The program will now end."
Pause
System

LetterAssignmentDone:

Return


' Local Subroutine

CheckTotalDiskSize:

' We are now going to check to see if the selected disk is larger than 2TB.

' Before going to this subroutine, make sure that the following variables are set:
' DiskID - The disk number to be checked
' SingleOrMulti$: Set to "SINGLE" if project is a single Windows image or "MULTI" if multiple images are to be made available for boot

' Initialize variables

Cls
Print "Check status of selected drive..."
DiskIDSearchString$ = "Disk" + Str$(DiskID) + " "
ff = FreeFile
Open "TEMP.BAT" For Output As #ff
Print #ff, "@echo off"
Print #ff, "(echo list disk"
Print #ff, "echo exit"
Print #ff, ") | diskpart > diskpart.txt"
Close #ff
Shell "TEMP.BAT"
If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"
ff = FreeFile
Open "diskpart.txt" For Input As #ff

' Init variables

AvailableSpaceString$ = ""
Units$ = ""
AvailableSpace = 0

Do
    Line Input #ff, ReadLine$

    If InStr(ReadLine$, DiskIDSearchString$) Then
        AvailableSpaceString$ = Mid$(ReadLine$, 28, 4)
        Units$ = Mid$(ReadLine$, 33, 2)
        Exit Do
    End If

Loop Until EOF(ff)

Close #ff

If _FileExists("diskpart.txt") Then Kill "diskpart.txt"
If Units$ = "MB" Then Multiplier = 1
If Units$ = "GB" Then Multiplier = 1024
If Units$ = "TB" Then Multiplier = 1048576

AvailableSpace = ((Val(AvailableSpaceString$)) * Multiplier)

' If the user has chosen to create a multi image boot disk, then a disk larger than 2TB is not valid under any
' circumstances. We will check the size of the disk now and if it is larger than 2TB and a multi image project
' was selected we will inform the user of this, abandon all operations, and return to the main menu.
'
' If the project is a single image type, then there are options available that would allow for the use of a
' disk larger than 2TB. In that case, we will present the user with these options.

' Note the value 2,097,152 below = The number of MB in 2 TB.

If (AvailableSpace > 2097152) Then

    AskForOverride:

    Cls
    Print "Do you want to set an MBR override? Type ";: Color 0, 10: Print "HELP";: Color 15: Print " for information about this option."
    Print

    ' Set Initial Value

    Override$ = ""

    Do
        Input "Do you wish to set an override"; Override$
    Loop While Override$ = ""

    If UCase$(Override$) = "HELP" Then
        Cls
        Print "You have selected a disk that is larger than 2TB."
        Print
        Print "The method used by this routine to make your disk bootable requires that the disk be initialized as MBR and not GPT."
        Print "This means that the disk size will be limited to 2TB. If you use a disk with more than 2TB capacity, then you will"
        Print "only be able to use 2TB of the space on that disk."
        Print
        Print "If you plan to boot this disk on only UEFI based systems, you can set an override. The program will then initialize"
        Print "the disk as GPT and allow the use of more than 2TB space. Note that setting this option only affects the disk when"
        Print "you completely initialize the disk. If you choose the option in the program to refresh your boot partitions, then"
        Print "the disk will remain MBR or GPT (whatever it currently is)."
        Pause
        GoTo AskForOverride
    End If

    YesOrNo Override$
    Override$ = YN$

    Select Case Override$
        Case "X"
            Print
            Color 14, 4: Print "Please provide a valid response.": Color 15
            Pause
            GoTo AskForOverride
        Case "Y"
            Exit Select
        Case "N"
            AvailableSpace = 2097152 - 2560
    End Select

End If

Return


' **************************************************************************
' * Create a bootable Windows ISO image that can include multiple editions *
' **************************************************************************

' NOTE: This code was originally a copy / edit of the code for the section "Inject Windows updates into one or more Windows ISO images"
' with the code for injecting updates removed. As such, there may be some variables or elements left over that don't seem to make sense.
' There may also be some variable names that elude to "updates" even though we are not actually injecting any updates.

MakeMultiBootImage:

' This routine will extract Windows editions from one or more ISO images and combine them into a single multi boot ISO image.

' Ask for source folder. Check to make sure folder contains ISO images. If it does, ask if all ISO images should be processed.
' For each image to be processed, we need to keep track of the image name to be processed. Likewise, we need to track source folder.

' Initialize variables

SourceFolder$ = ""
FileCount = 0
TotalFiles = 0
ReDim UpdateFlag(0) As String
ReDim FileArray(0) As String
ReDim FolderArray(0) As String ' A list of folders containing files to be processed. Note that the folder path stored here will end with a "\"

MMBIGetFolders:

Do
    Cls
    Print "Enter the path to one or more Windows ISO image files. These should be x64 only images. These images must include an"
    Print "install.wim file, ";: Color 0, 10: Print "NOT";: Color 15: Print " install.esd. ";: Color 0, 10: Print "DO NOT";: Color 15: Print " include a file name or extension."
    Print
    Line Input "Enter the path: ", SourceFolder$
Loop While SourceFolder$ = ""

CleanPath SourceFolder$
SourceFolder$ = Temp$ + "\"

' Verify that the path specified exists.

If Not (_DirExists(SourceFolder$)) Then
    Cls
    Print
    Color 14, 4
    Print "The location that you specified does not exist or is not valid."
    Color 15
    Pause
    Cls
    GoTo MMBIGetFolders
End If

' Perform a check to see if files with a .ISO extension exist in specified folder.
' We are going to call the FileTypeSearch subroutine for this. We will pass to it
' the path to search and the extension to search for. It will return to us the number
' of files with that extension in the variable called filecount and the name of
' each file in an array called FileArray$(x).

FileTypeSearch SourceFolder$, ".ISO", "N"

' FileTypeSearch returns number of ISO images found as NumberOfFiles and each of those files as TempArray$(x)

FileCount = NumberOfFiles
Cls

If FileCount = 0 Then
    Print
    Color 14, 4
    Print "No files with the .ISO extension were found.";
    Color 15
    Print " Please specify another folder."
    Pause
    Cls
    GoTo MMBIGetFolders
End If

' If we arrive here, then files with a .ISO extension were found at the location specified.
' FileCount holds the number of .ISO files found.

UpdateAll$ = "N" ' Set an initial value

MMBIUpdateAll:

' If there is only 1 file, then automatically set the UpdateAll$ flag to "Y", otherwise, ask if all files in folder should be updated.

Cls

If FileCount > 1 Then
    Print "Do you want to use at least one Windows edition from ";: Color 0, 10: Print "ALL";: Color 15: Print " of the files located here";: Input UpdateAll$
Else
    UpdateAll$ = "Y"
End If

YesOrNo UpdateAll$
UpdateAll$ = YN$
If UpdateAll$ = "X" Then GoTo MMBIUpdateAll

If UpdateAll$ = "Y" Then
    For x = 1 To FileCount
        TotalFiles = TotalFiles + 1

        ' Init variables

        ReDim _Preserve UpdateFlag(TotalFiles) As String
        ReDim _Preserve FileArray(TotalFiles) As String
        ReDim _Preserve FolderArray(TotalFiles) As String
        ReDim _Preserve FileSourceType(TotalFiles) As String

        UpdateFlag$(TotalFiles) = "Y"
        FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
        FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
        Cls
        Print "Please standby for a moment. Verifying the following image:"
        Print
        Color 10
        Print FileArray$(TotalFiles)
        Color 15
        Print
        Temp$ = FolderArray$(TotalFiles) + FileArray$(TotalFiles)
        DetermineArchitecture Temp$, 1
        Select Case ImageArchitecture$
            Case "x64"
                FileSourceType$(TotalFiles) = ImageArchitecture$
            Case "DUAL", "NONE", "x86"
                Cls
                Print
                Color 14, 4: Print "WARNING!";: Color 15: Print " An invalid file has been selected."
                Print
                Print "Check the following file to make sure that it is valid. It needs to contain an install.wim file, not INSTALL.ESD."
                Print "In addition, make sure that all files are valid x64 Windows images files."
                Print
                Print "Path: ";: Color 10: Print Left$(Temp$, ((_InStrRev(Temp$, "\"))) - 1): Color 15
                Print "File: ";: Color 10: Print Right$(Temp$, (Len(Temp$) - (_InStrRev(Temp$, "\")))): Color 15
                Pause
                ChDir ProgramStartDir$: GoTo BeginProgram
        End Select
    Next x
    GoTo MMBICheckForMoreFolders
End If

' We end up here if the user does NOT want to use every image in the selected location.
' In that case, we need to ask the user about each image file to see if it contains any
' Windows editions that the user wants to use.

For x = 1 To FileCount

    MMBIMarker1:

    Cls
    Print "Do you want to add any of the Windows editions in this file to your multi boot image?"
    Print
    Color 4: Print "Location:  ";: Color 10: Print Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
    Color 4: Print "File name: ";: Color 10: Print Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
    Color 15
    Print
    Input "Add editions from this file"; UpdateThisFile$
    YesOrNo UpdateThisFile$
    Select Case YN$
        Case "X"
            Print
            Color 14, 4
            Print "Please provide a valid response."
            Color 15
            Pause
            GoTo MMBIMarker1
        Case "Y"
            TotalFiles = TotalFiles + 1

            ' Check validity of selected files. Files should be valid x64 Windows image files.

            ' Init variables

            ReDim _Preserve UpdateFlag(TotalFiles) As String
            ReDim _Preserve FileArray(TotalFiles) As String
            ReDim _Preserve FolderArray(TotalFiles) As String
            ReDim _Preserve FileSourceType(TotalFiles) As String

            UpdateFlag$(TotalFiles) = "Y"
            FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
            FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
            Cls
            Print "Please standby for a moment. Verifying the following image:"
            Print
            Color 10
            Print FileArray$(TotalFiles)
            Color 15
            Print
            Temp$ = FolderArray$(TotalFiles) + FileArray$(TotalFiles)
            DetermineArchitecture Temp$, 1
            Select Case ImageArchitecture$
                Case "x64"
                    FileSourceType$(TotalFiles) = ImageArchitecture$
                Case "DUAL", "NONE", "x86"
                    Cls
                    Print
                    Color 14, 4: Print "WARNING!";: Color 15: Print " An invalid file has been selected."
                    Print
                    Print "Check the following file to make sure that it is valid. It needs to contain an install.wim file, not INSTALL.ESD."
                    Print "In addition, make sure that all files are valid x64 Windows image files."
                    Print
                    Print "Path: ";: Color 10: Print Left$(Temp$, ((_InStrRev(Temp$, "\"))) - 1): Color 15
                    Print "File: ";: Color 10: Print Right$(Temp$, (Len(Temp$) - (_InStrRev(Temp$, "\")))): Color 15
                    Pause
                    ChDir ProgramStartDir$: GoTo BeginProgram
            End Select
    End Select
Next x

MMBICheckForMoreFolders:

MoreFolders$ = "" ' Initial value
Cls
Input "Do you want to add another folder that has more ISO images"; MoreFolders$
YesOrNo MoreFolders$

Select Case YN$
    Case "X"
        Print
        Color 14, 4
        Print "Please provide a valid response."
        Color 15
        Pause
        GoTo MMBICheckForMoreFolders
    Case "Y"
        Cls
        GoTo MMBIGetFolders
End Select

' At this point, we have a list of all the folders and ISO image files that have Windows editions that are
' to be added to the multi boot image. First, we are going to verify that at least one file has been selected. If not,
' go back to the main menu.

If TotalFiles = 0 Then
    Cls
    Color 14, 4: Print "You have not selected any files to add to your image.";: Color 15: Print " We will now return to the main menu."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

MMBIGetIndexList:

ReDim IndexCount(TotalFiles) As Integer ' Initialize array

For IndexCountLoop = 1 To TotalFiles
    If UpdateFlag$(IndexCountLoop) = "N" Then GoTo MMBINoIndex

    MMBIGetMyIndexList:

    Cls
    Print "Enter the index number(s) for the Windows editions to be updated: "
    Print
    Color 4: Print "File Location: ";: Color 10: Print FolderArray$(IndexCountLoop)
    Color 4: Print "File Name    : ";: Color 10: Print FileArray$(IndexCountLoop)
    Color 15
    Print
    Print "> To view a list of available indices: Press ENTER."
    Print "> To view help for entering index numbers: Type HELP and press ENTER."
    Print "Otherwise, enter the index number(s) and press ENTER."
    Print
    Input "Enter index number(s), press ENTER, or type HELP and press ENTER: ", IndexRange$
    IndexRange$ = UCase$(IndexRange$)
    If IndexRange$ = "HELP" Then
        Cls
        Print "You can enter a single index number or multiple index numbers. To enter a contiguous range of index numbers,"
        Print "separate the numbers with a dash like this: 1-4. For non contiguous indices, separate them with a space like"
        Print "this: 1 3. You can also combine both methods like this: 1-3 5 7-9. Make sure to enter numbers from low to high."
        Print
        Print "Finally, if you want to add all editions of Windows, simply enter "; Chr$(34); "ALL"; Chr$(34); "."
        Pause
        GoTo MMBIGetMyIndexList
    End If
    If ((IndexRange$ <> "") And (IndexRange$ <> "ALL")) Then GoTo MMBIProcessRange
    SourcePath$ = FolderArray$(IndexCountLoop) + FileArray$(IndexCountLoop)

    Select Case IndexRange$
        Case ""
            Silent$ = "N"
        Case "ALL"
            Silent$ = "Y"
    End Select
    GoSub DisplayIndices2
    If IndexRange$ = "" Then GoTo MMBIGetMyIndexList
    If IndexRange$ = "ALL" Then
        Temp$ = ""
        GetNumberOfIndices
        Temp$ = _Trim$(Str$(NumberOfSingleIndices))
        IndexRange$ = "1-" + Temp$
        If IndexRange$ = "1-1" Then IndexRange$ = "1"
    End If
    Kill "Image_Info.txt"

    MMBIProcessRange:

    ProcessRangeOfNums IndexRange$, 1
    If ValidRange = 0 Then
        Color 14, 4
        Print "You did not enter a valid range of numbers"
        Color 15
        Pause
        GoTo MMBIGetMyIndexList
    End If

    ' We will now get image info and save it to a file called Image_Info.txt. We will parse that file to verify that the index
    ' selected is valid. If not, we will ask the user to choose a valid index.

    SourcePath$ = FolderArray$(IndexCountLoop) + FileArray$(IndexCountLoop)
    Print
    Print "Verifying indices."
    Print
    Print "Please standby..."
    Print
    GetWimInfo_Main SourcePath$, 1

    For x = 1 To TotalNumsInArray
        WimInfoFound = 0 ' Init Variable
        Open "Image_Info.txt" For Input As #1
        Do
            Line Input #1, WimInfo$
            If Len(WimInfo$) >= 9 Then
                If (Left$(WimInfo$, 7) = "Index :") And (Val(Right$(WimInfo$, (Len(WimInfo$) - 8))) = RangeArray(x)) Then
                    Line Input #1, WimInfo$
                    NameFromFile$ = Right$(WimInfo$, (Len(WimInfo$) - 7))
                    Line Input #1, WimInfo$
                    DescriptionFromFile$ = Right$(WimInfo$, (Len(WimInfo$) - 14))
                    WimInfoFound = 1
                End If
            End If

            MMBISkipToNextLine_Section1:

        Loop Until EOF(1)
        Close #1
        If WimInfoFound = 0 Then
            Cls
            Color 14, 4
            Print "Index"; RangeArray(x); "was not found."
            Print "Please supply a valid index number."
            Color 15
            Pause
            GoTo MMBIGetMyIndexList
        End If
        IndexCount(IndexCountLoop) = TotalNumsInArray

        ' For the index list, we are making an assumption that there will never be more than 100 indicies in an image.

        ReDim _Preserve IndexList(TotalFiles, 100) As Integer

        For y = 1 To TotalNumsInArray
            IndexList(IndexCountLoop, y) = RangeArray(y)
        Next y
    Next x
    Kill "Image_Info.txt"

    MMBINoIndex:

Next IndexCountLoop

' Now that we have a valid source directory and we know that there are ISO images
' located there, ask the user for the location where we should save the project

DestinationFolder$ = "" ' Set initial value

MMBIGetDestinationPath10:

Do
    Cls
    Print "Enter the path where the project will be created. We will use this location to save temporary files and we will also"
    Print "save the final ISO image file here."
    Print
    Line Input "Enter the path where the project should be created: ", DestinationFolder$
Loop While DestinationFolder$ = ""

CleanPath DestinationFolder$
DestinationFolder$ = Temp$ + "\"

' We don't want user to specify the root of a drive

If Len(DestinationFolder$) = 3 Then
    Cls
    Color 14, 4
    Print "Please do not specify the root directory of a drive."
    Color 15
    Pause
    GoTo MMBIGetDestinationPath10
End If

' Check to see if the destination specified is on a removable disk

Cls
Print "Performing a check to see if the destination you specified is a removable disk."
Print
Print "Please standby..."
DriveLetter$ = Left$(DestinationFolder$, 2)
RemovableDiskCheck DriveLetter$
DestinationIsRemovable = IsRemovable

Select Case DestinationIsRemovable
    Case 2
        Cls
        Color 14, 4: Print "This is not a valid disk.";: Color 15: Print " Please specify another location."
        Pause
        GoTo MMBIGetDestinationPath10
    Case 1
        Cls
        Print "The disk that you specified is a removable disk. ";: Color 14, 4: Print "Please specify a fixed disk.": Color 15
        Print
        Print "NOTE: Project must be created on a fixed disk due to limitations of some Microsoft utilities."
        Pause
        GoTo MMBIGetDestinationPath10
    Case 0
        ' if the returned value was a 0, no action is necessary. The program will continue normally.
End Select

' Verify that the path specified exists.

If Not (_DirExists(DestinationFolder$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(DestinationFolder$)) Then
        Cls
        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
        Print
        Print "Please recheck the path you have specified and try again."
        Pause
        GoTo MMBIGetDestinationPath10
    End If
End If

' If we have arrived here it means that the destination path already exists
' or we were able to create it successfully.

' Ask user what they want to name the final ISO image file

Cls
UserSelectedImageName$ = "" ' Set initial value
Print "If you would like to specify a name for the final ISO image file that this project will create, please do so now,"
Print "WITHOUT an extension. You can also simply press ENTER to use the default name of Windows.ISO."
Print
Print "Enter name ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension, or press ENTER: ";: Line Input "", UserSelectedImageName$

If UserSelectedImageName = "" Then
    UserSelectedImageName$ = "Windows.ISO"
Else
    UserSelectedImageName$ = UserSelectedImageName$ + ".ISO"
End If

' IMPORTANT: The count of files listed immediately below is the number of files of each type in the folders specified
' INCLUDING FILES THAT WILL NOT BE ADDED TO THE MULTI BOOT IMAGE.

' The next set of variables will hold the actual number of each image type to be processed

TotalImagesToUpdate = 0

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then TotalImagesToUpdate = TotalImagesToUpdate + IndexCount(x)
Next x

' If we reach this point, then the image specified by the user is valid.

' Before starting the process, verify that there are no leftover files sitting in the destination.

Cleanup DestinationFolder$
If CleanupSuccess = 0 Then ChDir ProgramStartDir$: GoTo BeginProgram

' Create the folders we need for the project.

MkDir DestinationFolder$ + "ISO_Files"
MkDir DestinationFolder$ + "WIM_x64"

' Export all the Windows editions to the WIM_x64 folder.

Cls
Print
Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Exporting All Windows Editions"
Print "[             ] Creating Base Image"
Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
Print "[             ] Creating Final ISO Image"

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        MountISO Temp$
        For y = 1 To IndexCount(x)
            SourceArcFlag$ = ""
            DestArcFlag$ = "WIM_x64"
            CurrentIndex$ = LTrim$(Str$(IndexList(x, y)))
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + MountedImageDriveLetter$ + SourceArcFlag$_
            + "\Sources\install.wim" + CHR$(34) + " /SourceIndex:" + CurrentIndex$ + " /DestinationImageFile:" + CHR$(34) + DestinationFolder$_
            + DestArcFlag$ + "\install.wim" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Next y

        ' The next command dismounts the ISO image since we are now done with it. The messages displayed by the process are
        ' not really helpful so we are going to hide those messages even if detailed status is selected by the user.

        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    End If
Next x

' Create the base image

' We need to know if all of the images to be updated are of the same architecture. This is because if any of the images come
' from a dual architecture image, even if they are all of the same architecture, the project is flagged as a dual architecture
' project. However, this results in no files from the other architecture type being copied to the project which in turn causes
' the ISO image to not operate properly. We will use the variable AllFilesAreSameArc to track whether or not all images are of
' the same architecture type or not.

Cls
Print
Print "[  COMPLETED  ] Exporting All Windows Editions"
Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Base Image"
Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
Print "[             ] Creating Final ISO Image"

AllFilesAreSameArc = 1
SingleImageTag = ""

' To ensure that DestinationFolder$ is always specified consistently without a trailing backslash, we will
' run it through the CleanPath routine.

CleanPath DestinationFolder$
DestinationFolder$ = Temp$

' Jump to the routine for creating a base image. Here we determine whether we need to run code for a Single or Dual Architecture project

MMBIProjectIsSingleArchitecture:

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        MountISO Temp$

        ' If an x64 folder exists, then even though the project is a single architecture type project, the source is a dual architecture source.
        ' This means that we need to copy the contents of the x64 or x86 folder to the root and not to the x64 or x86 folder.

        Select Case ExcludeAutounattend$
            Case "Y"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
                    + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd autounattend.xml > NUL"
            Case "N"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
                    + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd > NUL"
        End Select
        Shell Cmd$
        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Exit For
    End If
Next x

' When we arrive here, the base image for the project has been completed.

' We need a file called ei.cfg in each "sources" folder. When creating a Windows image that has multiple editions, this file is needed
' to prevent setup from simply installing the version of Windows that originally shipped on a machine without presenting a menu
' to allow you to choose the edition that you want to install. Note that this file is not needed for unattended installs since the
' autounattend.xml answer file will specify which version of Windows to install, but it does not hurt to have the file there.
'
' The following lines will check to see if an ei.cfg file is already present. If so, we will leave it alone in case it is configured
' differenty than what we are going to put in place, otherwise we will create the file.

If CreateEiCfg$ = "Y" Then
    Temp$ = DestinationFolder$ + "\ISO_Files\sources"

    If _DirExists(Temp$) Then
        If Not (_FileExists(Temp$ + "\ei.cfg")) Then
            Open (Temp$ + "\ei.cfg") For Output As #1
            Print #1, "[CHANNEL]"
            Print #1, "Retail"
            Close #1
        End If
    End If

End If

'MMBIDoneCreatingBaseImage:

' Moving the updated install.wim file(s) to the base image

Cls
Print
Print "[  COMPLETED  ] Exporting All Windows Editions"
Print "[  COMPLETED  ] Creating Base Image"
Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Moving Updated WIM File to Base Image and Syncing File Versions"
Print "[             ] Creating Final ISO Image"

Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
If _FileExists(Temp$) Then
    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
    Shell _Hide Cmd$
    Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "\WIM_x64\install.wim " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
    Shell _Hide Cmd$
End If

FinalImageName$ = DestinationFolder$ + "\" + UserSelectedImageName$

' Technical Note: OSCDIMG does not hide its output by simply redirecting to NUL. By using " > NUL 2>&1" we work around this.
' How this works: Standard output is going to NUL and standard error output (file descriptor 2) is being sent to standard output
' (file descriptor 1) so both error and normal output go to the same place.

Cls
Print
Print "[  COMPLETED  ] Exporting All Windows Editions"
Print "[  COMPLETED  ] Creating Base Image"
Print "[  COMPLETED  ] Moving Updated WIM File to Base Image and Syncing File Versions"
Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Final ISO Image"

' Clear the read-only, system, and hidden attributes from all source files

Cmd$ = "attrib -h -s -r " + Chr$(34) + DestinationFolder$ + "\ISO_Files\*.*" + Chr$(34) + " /s /d"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

' Create the final ISO image file

Do
    _Limit 10
    CurrentTime$ = Date$ + "," + Left$(Time$, 5)
    Select Case Right$(CurrentTime$, 8)
        Case "23:59:58", "23:59:59"
            Midnight = 1
        Case Else
            Midnight = 0
    End Select
Loop While Midnight = 1

Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -t" + CurrentTime$ + " -m -o -u2 -udfver102 -bootdata:2#p0,e,b" + CHR$(34) + DestinationFolder$ + "\ISO_Files\boot\etfsboot.com"_
+ CHR$(34) + "#pEF,e,b" + CHR$(34) + DestinationFolder$ + "\ISO_Files\efi\microsoft\boot\efisys.bin" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
+ "\ISO_Files" + CHR$(34) + " " + CHR$(34) + FinalImageName$ + CHR$(34) + " > NUL 2>&1"
Shell Chr$(34) + Cmd$ + Chr$(34)

Cls
Print
Print "[  COMPLETED  ] Exporting All Windows Editions"
Print "[  COMPLETED  ] Creating Base Image"
Print "[  COMPLETED  ] Moving Updated WIM File to Base Image and Syncing File Versions"
Print "[  COMPLETED  ] Creating Final ISO Image"

' Perform a quick cleanup
' Cleaning up previous stale DISM operations

Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /cleanup-wim"
Shell _Hide Cmd$

' Replacing the below command with a dism /cleanup-mountpoints in build 7.5.2.26. Commented lines can be removed once change is tested
' Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /online /cleanup-image /startcomponentcleanup"
' SHELL _HIDE Cmd$

Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /cleanup-mountpoints"
Shell _Hide Cmd$

' Cleanup folders

Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\WIM_x64" + Chr$(34) + " /s /q"
Shell _Hide Cmd$

' Clear keyboard buffer

_KeyClear

Print
Print "That's all!"
Print
Print "The final image file can be found here:"
Print
Color 4: Print "Location:  ";: Color 10: Print DestinationFolder$
Color 4: Print "File name: ";: Color 10: Print Mid$(FinalImageName$, (_InStrRev(FinalImageName$, "\") + 1)): Color 15
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' **************************************************************
' * Create a bootable ISO image from Windows files in a folder *
' **************************************************************

MakeBootDisk2:

' This routine will take a folder containing Windows files and create a bootable ISO image from those files.
' All files and folders must be present.

' Clear any pre-existing values for paths and filenames

MakeBootablePath$ = ""
DestinationFolder$ = ""
DestinationFileName$ = ""
VolumeName$ = ""

Do
    Cls
    Print "Enter the path to the folder with the Windows files to make into an ISO image: ";: Line Input "", MakeBootablePath$

Loop While MakeBootablePath$ = ""

CleanPath MakeBootablePath$
MakeBootablePath$ = Temp$
TempPath$ = MakeBootablePath$ + "\sources\install.wim"

' We cannot check for all files, but we are at least checking to see if an "INSTALL.WIM" is present at the specified location as a simple
' sanity check that the folder specified is likely valid.

If Not (_FileExists(TempPath$)) Then
    Print
    Color 14, 4: Print "That path is not valid.";: Color 15: Print " No install.wim file found at that location. Please try again."
    Pause
    GoTo MakeBootDisk2
End If

' If we reach this point, then the path provided exists and it contains an INSTALL.WIM file in the SOURCES folder.

' Determine if an autounattend.xml file exists in the dource folder

If _FileExists(MakeBootablePath$ + "\autounattend.xml") Then
    AnswerFilePresent$ = "Y"
Else
    AnswerFilePresent$ = "N"
End If

ISODestinationPath:

DestinationFolder$ = "" ' Set initial value

Do
    Cls
    Print "Enter the destination path. This is the path only ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " a file name: ";: Line Input "", DestinationFolder$
Loop While DestinationFolder$ = ""

CleanPath DestinationFolder$
DestinationFolder$ = Temp$

' Verify that the path specified exists.

If Not (_DirExists(DestinationFolder$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(DestinationFolder$)) Then
        Cls
        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
        Print
        Print "Please recheck the path you have specified and try again."
        Pause
        GoTo ISODestinationPath
    End If
End If

' If we have arrived here it means that the destination path already exists
' or we were able to create it successfully.

DestinationFileName$ = "" ' Set initial value

Do
    Cls
    Print "Enter the name of the file to create, ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension: ";: Line Input "", DestinationFileName$
Loop While DestinationFileName$ = ""

GetVolumeName1:

' Get the volume name for the ISO image

Cls
Line Input "Enter the volume name to give the ISO image or press Enter for none: ", VolumeName$

If Len(VolumeName$) > 32 Then
    Print
    Color 14, 4: Print "That volume name is invalid!";: Color 15: Print " The volume name is limited to 32 characters."
    Pause
    GoTo GetVolumeName1
End If

Cls

' Clear the read-only, system, and hidden attributes from all source files

Cmd$ = "attrib -h -s -r " + Chr$(34) + MakeBootablePath$ + "\*.*" + Chr$(34) + " /s /d"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

' Create the ISO image

Do
    _Limit 10
    CurrentTime$ = Date$ + "," + Left$(Time$, 5)
    Select Case Right$(CurrentTime$, 8)
        Case "23:59:58", "23:59:59"
            Midnight = 1
        Case Else
            Midnight = 0
    End Select
Loop While Midnight = 1

Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -t" + CurrentTime$ + " -m -o -u2 -udfver102 -l" + CHR$(34) + VolumeName$ + CHR$(34) + " -bootdata:2#p0,e,b" + CHR$(34)_
+ MakeBootablePath$ + "\boot\etfsboot.com" + CHR$(34) + "#pEF,e,b" + CHR$(34) + MakeBootablePath$ + "\efi\microsoft\boot\efisys.bin" + CHR$(34)_
+ " " + CHR$(34) + MakeBootablePath$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\" + DestinationFileName$ + ".iso" + CHR$(34) + " > NUL 2>&1"
Print "Creating the ISO image. Please standby..."
Shell Chr$(34) + Cmd$ + Chr$(34)

' Clear keyboard buffer

_KeyClear

Print
Print "Image created."

If AnswerFilePresent$ = "Y" Then
    Print
    Color 14, 4: Print "CAUTION!";: Color 15: Print " Your final image file contains an autounattend.xml answer file. If this image is booted, depending upon the"
    Print "configuration, it is possible that it could wipe the boot drive and / or other drives in the system on which it is"
    Print "being booted."
    Print
    Print "You may want to name this image and mark any media created from it to avoid accidental usage."
End If

Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' **************************************************
' * Reorganize the Contents of a Windows ISO Image *
' **************************************************

ChangeOrder:

' This routine will allow you to reorganize the order of Windows editions within an image so that they show up in the boot menu
' in the order that you want. You can also remove Windows editions from the image.

' Ask for source image file. Verify that it is a valid image.

GetImageToReorg:

' Initialize variable
SourceImage$ = ""

Cls
Line Input "Please enter the full path and file name of the image to reorganize: ", SourceImage$

CleanPath SourceImage$
SourceImage$ = Temp$

ReorgFileName$ = Mid$(SourceImage$, (_InStrRev(SourceImage$, "\") + 1))

If Not (_FileExists(SourceImage$)) Then
    Cls
    Color 14, 4: Print "No such file exists.";: Color 15: Print " Please specify a valid file."
    Pause
    GoTo GetImageToReorg
End If

ReorgSourcePath$ = Left$(SourceImage$, _InStrRev(SourceImage$, "\") - 1)
Print
Print "Standby while we verify the validity of this file..."

' Start by determining the architectre type of the image (either a single or dual architecture type).

DetermineArchitecture SourceImage$, 1

Select Case ImageArchitecture$
    Case "DUAL", "NONE", "x86"
        Cls
        Color 14, 4: Print "The image specified is not valid.";: Color 15: Print " Please specify a valid image."
        Pause
        GoTo ChangeOrder
End Select

'Init variables

IndexOrder$ = ""
SingleImageCount = 0

' Ask the user for the list of indices in the order in which they want them arranged.

ReorgSingle:

Do
    Cls
    Input "Enter the image order: ", IndexOrder$
Loop While IndexOrder$ = ""
ProcessRangeOfNums IndexOrder$, 0
If ValidRange = 0 Then
    Cls
    Color 14, 4: Print "You did not enter valid values.";: Color 15: Print " Enter a valid set of values.": Color 15
    GoTo ReorgSingle
End If
SingleImageCount = TotalNumsInArray
ReDim SingleArray(SingleImageCount) As Integer
For x = 1 To SingleImageCount
    SingleArray(x) = RangeArray(x)
Next x

' Verify that the index numbers specified are valid.

' We will now get image info and save it to a file called Image_Info.txt. We will parse that file to verify that the indices
' selected are valid. If not, we will ask the user to choose a valid index.

Print
Print "Verifying that the index number(s) supplied are valid."
Print
Print "Please standby..."
Print

' We are now going to get information regarding the WIM file(s) on the source. We will parse that information to determine what the highest numbered index is.
' This will allow us to check that no index numbers higher than valid have been supplied.

GetWimInfo_Main SourceImage$, 1

' Initialize variables

Highest_Single = 0

Open "Image_Info.txt" For Input As #1

Do
    Line Input #1, ReadLine$
    If InStr(ReadLine$, "Index :") Then
        Temp$ = (Right$(ReadLine$, (Len(ReadLine$) - _InStrRev(ReadLine$, ":"))))
    End If
Loop Until EOF(1)
Highest_Single = Val(Temp$)

' Close and delete the Image_Info.txt file since we are now done using it

Close #1
Kill "Image_Info.txt"

' Initialize variable

ValidRange = 1

For x = 1 To SingleImageCount
    If SingleArray(x) > Highest_Single Then ValidRange = 0
Next x

If ValidRange = 0 Then
    Cls
    Color 14, 4: Print "Invalid value detected!": Color 15
    Print
    Print "At least one of the index values specified is invalid. Please recheck the values and try again."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

ReorgGetDestination:

Cls
Print "Enter the destination path for the project. Please note that we will save the file using the same name that is has now."
Print "If you enter the same path where the source file is located, the old file will be replaced with the updated file."
Print
Line Input "Enter the destination for the project (path only, no file name or extension): ", Destination$
CleanPath Destination$
Destination$ = Temp$

If UCase$(Destination$) = UCase$(ReorgSourcePath$) Then
    Cls
    Print "You have specified the same location where the source file is located. This will cause that file to be"
    Print "replaced with the new file."
    Print
    Input "Do you want to replace the original file"; Temp$
    YesOrNo Temp$
    If YN$ = "Y" Then
        GoTo ReorgDestOkay
    Else GoTo ReorgGetDestination
    End If
End If

ReorgDestOkay:

If Not (_DirExists(Destination$)) Then
    Shell _Hide "md " + Chr$(34) + Destination$ + Chr$(34)
End If

If Not (_DirExists(Destination$)) Then
    Print
    Color 14, 4: Print "That folder does not exist and we were not able to create it.": Color 15
    Pause
    GoTo ReorgGetDestination
End If

' As a safety precaution, check to see if there are already folders by the name of those we are going to use for temporary
' storage of data. If there are, ask the user if it is okay to remove them.

'If _DirExists(Destination$ + "\x64") Then GoTo ReorgAssetExists
If _DirExists(Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34)) Then GoTo ReorgAssetExists
If _FileExists(Chr$(34) + Destination$ + "\install.wim" + Chr$(34)) Then GoTo ReorgAssetExists
GoTo ReorgCleanup

ReorgAssetExists:

Cls
Print "This routine creates a temporary folder called ISO_FILES and a file named install.wim in the destination folder."
Print "At least one of these already exists there."
Print
Input "Is it okay to erase these"; Temp$
YesOrNo Temp$
If YN$ = "Y" Then
    GoTo ReorgCleanup
Else
    GoTo ReorgGetDestination
End If

ReorgCleanup:

' Do a cleanup of the destination folder to get rid of any previously existing files

Cls
Print "Performing some quick housekeeping..."
Shell _Hide "rmdir " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34) + " /s /q"
Shell _Hide "del " + Chr$(34) + Destination$ + "\install.wim" + Chr$(34) + " /s /q"
Shell _Hide "md " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34)

' Mount the ISO image and copy the editions to our working folder in the proper order

Cls
Print "Exporting Windows editions..."
MountISO SourceImage$
ImageSourceDrive$ = MountedImageDriveLetter$

SRC$ = ImageSourceDrive$ + "\Sources\install.wim"
DST$ = Destination$ + "\install.wim"
For x = 1 To SingleImageCount
    Locate 3, 1: Print "Exporting image"; x; "of"; SingleImageCount
    IDX$ = LTrim$(Str$(SingleArray(x)))
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Export-Image /SourceImageFile:" + SRC$ + " /SourceIndex:" + IDX$ + " /DestinationImageFile:" + Chr$(34) + DST$ + Chr$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
Next x

Cls
Print "Reorganizing the image in the order you specified."

' Copy the files needed to create the base image.

Cmd$ = "robocopy " + Chr$(34) + ImageSourceDrive$ + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xf install.wim"
Shell Cmd$

If _FileExists(Destination$ + "\install.wim") Then
    Cmd$ = "move /Y " + Chr$(34) + Destination$ + "\install.wim" + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
    Shell _Hide Cmd$
End If

' Save the volume information from the original image. We will run a "dir" against the source image and parse it for the volume label. We will capture that
' volume label and apply it to the new image that we are creating.

Cmd$ = "dir " + ImageSourceDrive$ + " > ISOImageInfo.txt"
Shell Cmd$
Open "IsoImageInfo.txt" For Input As #1
Line Input #1, ImageInfo$
Close #1

If InStr(ImageInfo$, "has no label") Then
    VolumeLabel$ = ""
Else
    VolumeLabel$ = Right$(ImageInfo$, (Len(ImageInfo$) - 22))
End If

Cls

' Clear the read-only, system, and hidden attributes from all source files

Cmd$ = "attrib -h -s -r " + Chr$(34) + DestinationFolder$ + "\ISO_Files\*.*" + Chr$(34) + " /s /d"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

' Dismount the original image

Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourceImage$ + "'" + Chr$(34) + Chr$(34)
Shell _Hide Cmd$

' Create the final ISO image

Print "Creating the final ISO image. Please standby..."

Do
    _Limit 10
    CurrentTime$ = Date$ + "," + Left$(Time$, 5)
    Select Case Right$(CurrentTime$, 8)
        Case "23:59:58", "23:59:59"
            Midnight = 1
        Case Else
            Midnight = 0
    End Select
Loop While Midnight = 1

Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -t" + CurrentTime$ + " -m -o -u2 -udfver102 -l" + CHR$(34) + VolumeLabel$ + CHR$(34) + " -bootdata:2#p0,e,b" + CHR$(34)_
+ Destination$ + "\ISO_Files\boot\etfsboot.com" + CHR$(34) + "#pEF,e,b" + CHR$(34) + Destination$ + "\ISO_Files\efi\microsoft\boot\efisys.bin"_
+ CHR$(34) + " " + CHR$(34) + Destination$ + "\ISO_Files" + CHR$(34) + " " + CHR$(34) + Destination$ + "\" + ReorgFileName$ + CHR$(34) + " > NUL 2>&1"
Shell Chr$(34) + Cmd$ + Chr$(34)

' Project is done. Cleanup the temporary files.

Cls
Print "Removing the temporary files used to create the new image..."
Shell _Hide "rmdir " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34) + " /s /q"
Shell _Hide "del " + Chr$(34) + Destination$ + "\install.wim" + Chr$(34) + " /s /q"

' Clear keyboard buffer

_KeyClear

' Inform the user that the project is done and then return to the main menu.

Cls
Print "Done!"
Print
Print "The updated file can be found here:"
Print
Color 10: Print Destination$: Color 15
Print
Print "The file name is the same as the original file: ";: Color 10: Print ReorgFileName$: Color 15
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' ***********************************************************************
' * Convert between an ESD and WIM either standalone or in an ISO image *
' ***********************************************************************

ConvertEsdOrWim:

_ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " - Convert Between ESD and WIM"

' Set initial values for variables

ImageType$ = ""
OriginalImageType$ = ""
Source$ = ""

Do
    Cls
    Print "Please enter the location of the Windows ISO image, Install.ESD, or Install.WIM file to be converted. You must"
    Print "specify the full path including the filename and extension."
    Print
    Line Input "Enter the path and file name: ", Source$
Loop While Source$ = ""

CleanPath Source$
Source$ = Temp$

' The file specified must be a .ISO, .ESD, or .WIM file. Check to make sure that the user specified one of these file types.

Temp$ = UCase$(Right$(Source$, 4))
Select Case Temp$
    Case ".ISO", ".ESD", ".WIM"
        Exit Select
    Case Else
        Cls
        Color 14, 4: Print "Invalid entry!";: Color 15: Print " Please enter the full path including the file name and extension. The file specified must be a"
        Print ".ISO, .ESD, or .WIM file."
        Pause
        GoTo ConvertEsdOrWim
End Select

'Check to make sure that the file specified exists.

If Not _FileExists(Source$) Then
    Cls
    Color 14, 4: Print "Invalid entry!";: Color 15: Print " Please specify a valid path and file name."
    Pause
    GoTo ConvertEsdOrWim
End If

' The user has specified a valid file name. Set a variable to hold the type of file to be processed.

OriginalImageType$ = Right$(Temp$, 3)

GetConversionDestinationPath:

DestinationFolder$ = "" ' Set an initial value for the destination path

Do
    Cls
    Print "Enter the path where the project will be created. This is where all the temporary files will be stored and we will"
    Print "save the final image file here as well."
    Print
    Line Input "Enter the path where the project should be created: ", DestinationFolder$
Loop While DestinationFolder$ = ""

CleanPath DestinationFolder$
DestinationFolder$ = Temp$

' If the file type specified is a .ISO image then we will need a name to give the final ISO image that will be created.

If OriginalImageType$ = "ISO" Then
    FinalImageName$ = "" ' Set initial value for the final image name
    Do
        Cls
        Line Input "Enter a name for the final image (without an extension): ", FinalImageName$
    Loop While FinalImageName$ = ""
    FinalImageName$ = DestinationFolder$ + "\" + FinalImageName$ + ".iso"
End If

' Check to see if the destination specified is on a removable disk

Cls
Print "Performing a check to see if the destination you specified is a removable disk."
Print
Print "Please standby..."
DriveLetter$ = Left$(DestinationFolder$, 2)
RemovableDiskCheck DriveLetter$
DestinationIsRemovable = IsRemovable

Select Case DestinationIsRemovable
    Case 2
        Cls
        Color 14, 4: Print "This is not a valid disk.";: Color 15: Print " Please specify another location."
        Pause
        GoTo GetConversionDestinationPath
    Case 1
        Cls
        Print "The disk that you specified is a removable disk. ";: Color 14, 4: Print "Please specify a fixed disk.": Color 15
        Print
        Print "NOTE: Project must be created on a fixed disk due to limitations of some Microsoft utilities."
        Pause
        GoTo GetConversionDestinationPath
    Case 0
        ' if the returned value was a 0, no action is necessary. The program will continue normally.
End Select

' Verify that the path specified exists.

If Not (_DirExists(DestinationFolder$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(DestinationFolder$)) Then
        Cls
        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
        Print
        Print "Please recheck the path you have specified and try again."
        Pause
        GoTo GetConversionDestinationPath
    End If
End If

' If we have arrived here it means that the destination path already exists
' or we were able to create it successfully.

' If the file type being processed is an ISO image, then we should perform a cleanup on the destination to make sure that
' files from a previous run are not still present. For ESD and WIM files this is not necessary. However, for an ESD or WIM
' we need to set the variable ImageType$ to indicate the type of file being converted.

Select Case OriginalImageType$
    Case "ISO"
        GoTo SetupForIsoImage
    Case "ESD", "WIM"
        ImageType$ = OriginalImageType$
        GoTo ProcessEsdOrWim
End Select

' Perform a cleanup of the destination path and prepare for processing an ISO image file

SetupForIsoImage:

Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
Cls
Print "Copying files from the ISO image and determining how many Windows editions are present. This may take a little while."
Print
Print "Please standby..."

MountISO Source$

If _DirExists(MountedImageDriveLetter$ + "\x64") Then
    Cls
    Print "The image you have specified is a dual architecture image. This x64 only edition of the program will not support"
    Print "dual architecture images."
    Pause

    ' Unmount the image and restart the program

    Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + Source$ + "'" + Chr$(34) + Chr$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

' Verify that the ISO image specified contains either an install.wim or an install.esd file
' Note that the original image type as specified by the user (ISO, ESD, or WIM) is stored in
' the variable called OriginalImageType$. While processing an ISO image, we determine if the
' ISO image contains an Install.ESD or an Install.WIM file. The variable ImageType$ holds the
' type of image (ESD or WIM) within the ISO image.

If _FileExists(MountedImageDriveLetter$ + "\Sources\install.esd") Then
    ImageType$ = "ESD"
ElseIf _FileExists(MountedImageDriveLetter$ + "\Sources\install.wim") Then
    ImageType$ = "WIM"
Else
    ImageType$ = "NONE"
End If

' If neither an Install.ESD or Install.WIM are found, cancel this routine and go back to the main menu

If ImageType$ = "NONE" Then
    Cls
    Color 14, 4: Print "Error!";: Color 15: Print " Neither an Install.WIM nor an Install.ESD file were found."
    Pause

    ' Since no install.wim or install.esd file was found, we will cleanup the working directory and return to the main menu

    Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /s /q"
    Shell _Hide Cmd$
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

' Continue here if we found either an Install.ESD or Install.WIM file

Shell "md " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34)
Cmd$ = "robocopy " + MountedImageDriveLetter$ + "\ " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /mir /a-:rsh /njh /njs"
Shell _Hide Cmd$

' The next command dismounts the ISO image since we are now done with it. The messages displayed by the process are
' not really helpful so we are going to hide those messages even if detailed status is selected by the user.

Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + Source$ + "'" + Chr$(34) + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

ProcessEsdOrWim:

' Determine how many editions of Windows are in the source

Select Case OriginalImageType$
    Case "ISO"
        ImagePath$ = DestinationFolder$ + "\ISO_Files\sources\install." + ImageType$
    Case "ESD", "WIM"
        ImagePath$ = Source$
End Select

Cmd$ = "dism /get-wiminfo /wimfile:" + Chr$(34) + ImagePath$ + Chr$(34) + " > ImageInfo.txt"
Shell Cmd$

ff = FreeFile
Open "ImageInfo.txt" For Input As #ff
Temp$ = _ReadFile$("ImageInfo.txt")
Close #ff

If _FileExists("ImageInfo.txt") Then
    Kill "ImageInfo.txt"
End If

' Set initial variable values

IndexCount = 0
SearchPosition = 1

' Search the output of the file to which we saved image info to determine how many indices are present

Do
    SearchPosition = InStr(SearchPosition + 1, Temp$, "Index :")
    If SearchPosition > 0 Then
        IndexCount = IndexCount + 1
    End If
Loop Until SearchPosition = 0

If IndexCount = 0 Then
    Cls
    Color 14, 4: Print "Error!";: Color 15: Print " We found no Windows editions in the image file."
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

x = 0 ' Set an initial value. This variable is used as a simple counter.

Do
    Cls
    x = x + 1

    ' Strip the leading space from the index number and index count by saving it as string variable.

    Index$ = LTrim$(Str$(x))
    IndexCountString$ = LTrim$(Str$(IndexCount))

    Select Case ImageType$
        Case "ESD"
            Print "Converting an ESD file into a WIM file"
        Case "WIM"
            Print "Converting a WIM file into an ESD file"
    End Select

    Print "Exporting Index #"; Index$; " of "; IndexCountString$

    If x = 1 Then
        If IndexCount > 1 Then
            Print
            Print "TIP: This first image will take much longer to export than the images that follow it."
        End If
    End If

    Select Case ImageType$
        Case "ESD"
            _ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " - Convert ESD to WIM - Currently Servicing Index #" + Index$ + " of " + IndexCountString$

            If OriginalImageType$ = "ISO" Then
Cmd$ = chr$(34) + "DISM /Export-Image /SourceImageFile:" + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources\install.esd"_
             + Chr$(34) + " /SourceIndex:" + Index$ + " /DestinationImageFile:" + Chr$(34) + DestinationFolder$_
              + "\ISO_Files\Sources\install.wim" + Chr$(34) + " /Compress:Max /CheckIntegrity" + Chr$(34) +chr$(34)
                Shell Cmd$
            Else
            Cmd$ = chr$(34) + "DISM /Export-Image /SourceImageFile:" + Chr$(34) + source$_
             + Chr$(34) + " /SourceIndex:" + Index$ + " /DestinationImageFile:" + Chr$(34) + DestinationFolder$_
              + "\install.wim" + Chr$(34) + " /Compress:Max /CheckIntegrity" + Chr$(34) +chr$(34)
                Shell Cmd$
            End If

        Case "WIM"
            _ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " - Convert WIM to ESD - Currently Servicing Index #" + Index$ + " of " + IndexCountString$

            If OriginalImageType$ = "ISO" Then
Cmd$ = chr$(34) + "DISM /Export-Image /SourceImageFile:" + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources\install.wim"_
             + Chr$(34) + " /SourceIndex:" + Index$ + " /DestinationImageFile:" + Chr$(34) + DestinationFolder$_
              + "\ISO_Files\Sources\install.esd" + Chr$(34) + " /Compress:Recovery /CheckIntegrity"+ Chr$(34) + Chr$(34)
                Shell Cmd$
            Else
            Cmd$ = chr$(34) + "DISM /Export-Image /SourceImageFile:" + Chr$(34) + source$_
             + Chr$(34) + " /SourceIndex:" + Index$ + " /DestinationImageFile:" + Chr$(34) + DestinationFolder$_
              + "\install.esd" + Chr$(34) + " /Compress:Recovery /CheckIntegrity"+ Chr$(34) + Chr$(34)
                Shell Cmd$
            End If

    End Select

Loop Until x = IndexCount

ExportsDone:

' The ESD or WIM file has now been converted. If the original file specified was an ISO image, then delete the old file
' since it has been converted and is no longer needed.If the original file was an ESD or WIM, then there is no need to
' create an ISO image or perform a cleanup.


If OriginalImageType$ = "ISO" Then

    Select Case ImageType$
        Case "ESD"
            Kill DestinationFolder$ + "\ISO_Files\Sources\install.esd"
        Case "WIM"
            Kill DestinationFolder$ + "\ISO_Files\Sources\install.wim"
    End Select

    Cls
    Print "Creating final image."
    Print
    Print "Please standby..."

    _ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " - Convert Between ESD and WIM - Creating Final Image"

    ' Unlike the routines where we alter the time of the files, since we are simply converting ESD to WIM in this routine, we will
    ' keep the original timestamps and won't alter them here.

Cmd$ = Chr$(34) + OSCDIMGLocation$ + Chr$(34) + " -m -o -u2 -udfver102 -bootdata:2#p0,e,b" + Chr$(34) + DestinationFolder$_
 + "\ISO_Files\boot\etfsboot.com" + Chr$(34) + "#pEF,e,b" + Chr$(34) + DestinationFolder$ + "\ISO_Files\efi\microsoft\boot\efisys.bin"_
 + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " " + Chr$(34) + FinalImageName$ + Chr$(34)

    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

    ' We are done with the ISO_Files folder that we used to temporarily store the Windows image files. Delete it.

    _ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " - Convert Between ESD and WIM"

    Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /s /q"
    Shell _Hide Cmd$
Else
    _ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " - Convert Between ESD and WIM"
End If

' Clear keyboard buffer

_KeyClear

' We are finished. Let the user know that we are done.

Cls
Print
Color 0, 10: Print "All processes have been completed.": Color 15
Print

Select Case OriginalImageType$
    Case "ISO"
        Print "You have converted an ISO image that contained an Install."; ImageType$; " into an image with an ";
        If ImageType$ = "ESD" Then Print "Install.WIM file."
        If ImageType$ = "WIM" Then Print "Install.ESD file."
        Print "The updated file can be found here:"
        Print
        Color 10: Print FinalImageName$: Color 15
    Case "ESD"
        Print "You have converted an ESD image into a WIM image."
        Print "The updated file can be found here:"
        Print
        Color 10: Print DestinationFolder$; "\install.wim": Color 15
    Case "WIM"
        Print "You have converted a WIM image into an ESD image."
        Print "The updated file can be found here:"
        Print
        Color 10: Print DestinationFolder$; "\install.esd": Color 15
End Select

Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' *********************************************************************************************************
' * Get image info - display basic info for each edition in an ISO image and display Windows build number *
' *********************************************************************************************************

' NOTE: This routine has some fairly major differences from the equivalent routine in the Dual Architecture
' version of this program. This is because we don't need to deal with dual architecture images in this edition
' of the program which simplies things considerably. Most notably, the build number is displayed before the
' details of individual Windows editions whereas in the dual architecture edition of this program we display
' that at the end.

GetWimInfo:

Do
    Cls
    Print "Enter the full path to the ISO image from which to get information."
    Line Input "Include the file name and extension: ", SourcePath$
Loop While SourcePath$ = ""

CleanPath SourcePath$
SourcePath$ = Temp$

If Not (_FileExists(SourcePath$)) Then
    Cls
    Color 14, 4: Print "No such file name could be found.";: Color 15: Print " Please try again."
    Pause
    GoTo GetWimInfo
End If

GetWimInfo_Main SourcePath$, 0

' Just in case any keys were pressed during the gathering of the WIM info, clear the keyboard buffer

_KeyClear

' Display Image_Info.txt which lists all of the indicies, then delete it if user no longer needs it.

Cls
DisplayFile "Image_Info.txt"

AskToSaveWimInfo1:
Cls
Print "A copy of the information just displayed can be saved to a file named Image_Info.txt in the same location"
Print "where this program is located so that you can refer to it if needed."
Print
Input "Do you want to save that file now"; Temp$
YesOrNo Temp$
Temp$ = YN$

If Temp$ = "X" Then
    Print
    Color 14, 4: Print "Please respond with a valid answer.": Color 15
    GoTo AskToSaveWimInfo1
End If

If Temp$ = "N" Then
    Kill "Image_Info.txt"
End If

If Temp$ = "Y" Then
    Shell Chr$(34) + "move Image_Info.txt " + Chr$(34) + ProgramStartDir$ + Chr$(34) + " > NUL" + Chr$(34)
    Cls
    Print "The file has been saved as:"
    Print
    Color 10: Print ProgramStartDir$; "Image_Info.txt": Color 15
    Pause
End If

ChDir ProgramStartDir$: GoTo BeginProgram


' ********************************************************************
' * Modify the NAME and DESCRIPTION values for entries in a WIM file *
' ********************************************************************

NameAndDescription:

' Initialize variables

ArchitectureChoice$ = ""

Do
    Cls
    Print "Enter the full path to the ISO image file that you want to work with."
    Line Input "Include the file name and extension: ", SourcePath$

Loop While SourcePath$ = ""

CleanPath SourcePath$
SourcePath$ = Temp$

If Not (_FileExists(SourcePath$)) Then
    Cls
    Color 14, 4: Print "No such file name could be found.";: Color 15: Print " Please try again."
    Pause
    GoTo NameAndDescription
End If

' Determine if the file specified holds a dual architecture installation
' Unlike the routine for creating a multiboot disk, in this case we only
' need to know if the architecture is dual or single so the only values
' we use for ProjectArchitecture$ are DUAL or SINGLE.

Cls
Print "Checking to see if the image specified is valid."
Print
MountISO SourcePath$
Temp$ = MountedImageDriveLetter$ + "\x64"

If _DirExists(Temp$) Then
    ProjectArchitecture$ = "DUAL"
Else
    ProjectArchitecture$ = "SINGLE"
End If

Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourcePath$ + "'" + Chr$(34) + Chr$(34) + " > NUL"
Shell Cmd$

If ProjectArchitecture = "DUAL" Then
    Cls
    Print
    Color 14, 4: Print "WARNING!";: Color 15: Print " An invalid file has been selected."
    Print
    Print "Make sure that that the file is a valid x64 Windows image file."
    Print
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

Do
    Cls
    Print "Note that you can save your updated image to the same location where the original is located."
    Print "You can even use the same file name if you want to update that file in its current location."
    Print
    Line Input "Enter the destination path without a file name or extension: ", DestinationFolder$
Loop While DestinationFolder$ = ""

CleanPath DestinationFolder$
DestinationFolder$ = Temp$ + "\"

Do
    Cls
    Print "Enter the name of the ISO image file to create ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension: ";: Line Input "", OutputFileName$
Loop While OutputFileName$ = ""

' Get the volume name for the ISO image

GetVolumeName3:

Cls
Print "Enter the volume name to give the ISO image or press ENTER for none (32 characters maximum)."
Print
Input "Volume name or ENTER for none: ", VolumeName$

If Len(VolumeName$) > 32 Then
    Print
    Color 14, 4: Print "That volume name is invalid!";: Color 15: Print " The volume name is limited to 32 characters."
    Pause
    GoTo GetVolumeName3
End If

Cls
Print
Print "Please be patient. We will ask for more information after staging the image to the project folder."
Print "This may take a little while."
Print
Print "*******************************"
Print "* Cleaning the project folder *"
Print "*******************************"
Print
Shell "rmdir /s /q " + Chr$(34) + DestinationFolder$ + "ISO_Files" + Chr$(34) + " > NUL 2>&1"
Shell "md " + Chr$(34) + DestinationFolder$ + "ISO_Files"
Print "*********************************"
Print "* Mounting the source ISO image *"
Print "*********************************"
Print
MountISO SourcePath$
Print "*************************************************************"
Print "* Copying files from the source image to the project folder *"
Print "*************************************************************"
Shell "robocopy " + MountedImageDriveLetter$ + "\" + " " + Chr$(34) + DestinationFolder$ + "ISO_Files" + Chr$(34) + " /mir /njh /njs /nfl /ndl /a-:r"
Print "*************************"
Print "* Dismounting the image *"
Print "*************************"
Print
Shell "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourcePath$ + "'" + Chr$(34) + Chr$(34) + " > NUL"

PickAnIndex:

ChosenIndex = 0 ' Set initial value

ArchitectureChoice$ = ""

Do
    Cls
    Print "> To view a list of available indices along with the current NAME and DESCRIPTION for each, press ENTER."
    Print "Otherwise, enter the index number for the edition to be updated and press ENTER."
    Print
    Input "Index Number: ", ChosenIndex
    If ChosenIndex = 0 Then
        Cls
        Print "Preparing to display a list of Windows editions and the associated index numbers."
        Print
        Print "Please standby..."
        Print
        Silent$ = "N"
        GoSub DisplayIndices2

        AskToSaveWimInfo5:

        Cls
        Print "A copy of the information just displayed can be saved to a file named Image_Info.txt in the same location"
        Print "where this program is located so that you can refer to it if needed."
        Print
        Input "Do you want to save that file now"; Temp$
        YesOrNo Temp$
        Temp$ = YN$
        If Temp$ = "X" Then
            Print
            Color 14, 4: Print "Please respond with a valid answer.": Color 15
            GoTo AskToSaveWimInfo5
        End If
        If Temp$ = "N" Then
            Kill "Image_Info.txt"
        End If
        If Temp$ = "Y" Then
            Shell Chr$(34) + "move Image_Info.txt " + Chr$(34) + ProgramStartDir$ + Chr$(34) + " > NUL" + Chr$(34)
        End If
        Cls
    End If
Loop Until ChosenIndex <> 0

IndexString$ = Str$(ChosenIndex)
IndexString$ = Right$(IndexString$, ((Len(IndexString$) - 1)))

Do
    Cls
    Input "Enter the NAME to assign to this entry: ", EditionName$
Loop While EditionName$ = ""

Do
    Cls
    Input "Enter the DESCRIPTION to assign to this entry: ", Description$
Loop While Description$ = ""

Cls
Print "*************************************************"
Print "* Updating the NAME and DESCRIPTION information *"
Print "*************************************************"
Print
Cmd$ = CHR$(34) + IMAGEXLocation$ + CHR$(34) + " /info " + CHR$(34) + DestinationFolder$ + "iso_files" + ArchitectureChoice$ + "\sources\install.wim" + CHR$(34)_
+ " " + IndexString$ + " " + CHR$(34) + EditionName$ + CHR$(34) + " " + CHR$(34) + Description$ + CHR$(34) + " /check > NUL"+chr$(34)
Shell Chr$(34) + Cmd$ + Chr$(34)
Print

Do
    Cls
    Print "Do you want to update another Windows edition in the ";: Color 0, 10: Print "SAME";: Color 15: Print " image file";
    Input ""; Another$
    YesOrNo Another$
Loop While YN$ = "X"

If YN$ = "Y" Then
    GoTo PickAnIndex
End If

Cls
Print "********************************"
Print "* Creating the final ISO image *"
Print "********************************"
Print
Do
    _Limit 10
    CurrentTime$ = Date$ + "," + Left$(Time$, 5)
    Select Case Right$(CurrentTime$, 8)
        Case "23:59:58", "23:59:59"
            Midnight = 1
        Case Else
            Midnight = 0
    End Select
Loop While Midnight = 1

Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -t" + CurrentTime$ + " -m -o -u2 -udfver102 -l" + CHR$(34) + VolumeName$ + CHR$(34) + " -bootdata:2#p0,e,b" + CHR$(34)_
+ DestinationFolder$ + "ISO_Files\boot\etfsboot.com" + CHR$(34) + "#pEF,e,b" + CHR$(34) + DestinationFolder$ + "ISO_Files\efi\microsoft\boot\efisys.bin"_
+ CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "ISO_Files" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + OutputFileName$ + ".ISO" + CHR$(34) + " > NUL 2>&1"
Shell Chr$(34) + Cmd$ + Chr$(34)
Print "*******************************"
Print "* Cleaning up temporary files *"
Print "*******************************"
Print
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "ISO_Files" + Chr$(34) + " /s /q"
Shell _Hide Cmd$

' Clear the keyboard buffer

_KeyClear

Print "*********************"
Print "* Project completed *"
Print "*********************"
Print
Print "The final image file can be found here:"
Print
Color 4: Print "Location:  ";: Color 10: Print DestinationFolder$
Color 4: Print "File name: ";: Color 10: Print OutputFileName$; ".ISO": Color 15
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' ***********************************
' * Export drivers from this system *
' ***********************************

ExportDrivers:

ExportFolder$ = "" ' Set initial value

Do
    Cls
    Line Input "Enter the full path to the location where you want the drivers to be exported: ", ExportFolder$
Loop While ExportFolder$ = ""

CleanPath ExportFolder$
ExportFolder$ = Temp$

' We don't want user to specify the root of a drive

If Len(ExportFolder$) = 3 Then
    Cls
    Color 14, 4: Print "Please do not specify the root directory of a drive.": Color 15
    Pause
    GoTo ExportDrivers
End If

' Verify that the path specified exists.

If Not (_DirExists(ExportFolder$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cmd$ = "md " + Chr$(34) + ExportFolder$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(ExportFolder$)) Then
        Cls
        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
        Print
        Print "Please recheck the path you have specified and try again."
        Pause
        GoTo ExportDrivers
    End If
End If

' If we have arrived here it means that the path specified already exists
' or we were able to create it successfully.

Cls
Print "Drivers are now being exported."
Print "Please be patient since this process may take a while..."
Cmd$ = Chr$(34) + "%SystemRoot%\system32\pnputil.exe" + Chr$(34) + " /export-driver * " + Chr$(34) + ExportFolder$ + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

ff = FreeFile
Open ExportFolder$ + "\Install_Drivers.bat" For Output As #ff
Print #ff, ":: Pass a 1 to this file if no reboot is desired. If anything other than a 1 is passed, a reboot"
Print #ff, ":: will performed after drivers are installed."
Print #ff, ""
Print #ff, "@echo off"
Print #ff, ""
Print #ff, ":: This next block checks to see if any parameter was passed to this batch file."
Print #ff, ""
Print #ff, "if X%1==X ("
Print #ff, "set NOREBOOT=0"
Print #ff, ") ELSE ("
Print #ff, "set NOREBOOT=%1"
Print #ff, ")"
Print #ff, ""
Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ":: Change to the directory where the batch file is run from ::"
Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ""
Print #ff, "cd /d %~dp0"
Print #ff, ""
Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ":: Check to see if this batch file is being run as Administrator. If it is not, then rerun the batch file ::"
Print #ff, ":: automatically as admin and terminate the initial instance of the batch file.                           ::"
Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ""
Print #ff, "(Fsutil Dirty Query %SystemDrive%>Nul)||(PowerShell start "; Chr$(34); Chr$(34); Chr$(34); "%~f0"; Chr$(34); Chr$(34); Chr$(34); " -verb RunAs & Exit /B)"
Print #ff, ""
Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ":: End Routine to check if being run as Admin ::"
Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ""
Print #ff, ":::::::::::::::::::::"
Print #ff, ":: Install drivers ::"
Print #ff, ":::::::::::::::::::::"
Print #ff, ""
Print #ff, "cls"
Print #ff, "echo *******************************"
Print #ff, "echo * Drivers Are Being Installed *"
Print #ff, "echo *******************************"
Print #ff, "echo."
Print #ff, "echo Please be patient since this process may take a while..."
Print #ff, ""
Print #ff, "pnputil /add-driver *.inf /subdirs /install > NUL"
Print #ff, ""
Print #ff, "if %NOREBOOT% NEQ 1 shutdown /r /t 10 > NUL"
Close #ff

' Clear the keyboard buffer

_KeyClear

Cls
Print "Drivers have been exported to the following location:"
Print
Color 10: Print ExportFolder$: Color 15
Print
Print "A file called ";: Color 10: Print "Install_Drivers.bat";: Color 15: Print " has also been created to install the drivers. Simply run that batch file to install"
Print "all the drivers that were just exported. Note that a reboot may be needed after installing the drivers."
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' ******************************************
' * Expand drivers supplied in a .CAB file *
' ******************************************

ExpandDrivers:

' This routine will expand drivers that are distributed in .CAB files. Once expanded, Menu Item #5 can
' then be used to inject these drivers into Windows ISO images. Note that not all .CAB files can be opened by the
' Windows EXPAND utility. Files that cannot be opened will simply be copied to the destination.

SourceFolder$ = "" ' Set initial value

Do
    Cls
    Print "Enter the path to the .CAB files that you wish to expand. Enter the path only with no file name or extension."
    Print
    Line Input "Enter the path to the drivers that are in .CAB files: ", SourceFolder$
Loop While SourceFolder$ = ""

CleanPath SourceFolder$
SourceFolder$ = Temp$ + "\"

' Verify that the path specified exists.

If Not (_DirExists(SourceFolder$)) Then
    Cls
    Color 14, 4: Print "The location that you specified does not exist or is not valid.": Color 15
    Pause
    GoTo ExpandDrivers
End If

' Perform a check to see if files with a .CAB extension exist in specified folder.
' We are going to call the FileTypeSearch subroutine for this. We will pass to it
' the path to search and the extension to search for. It will return to us the number
' of files with that extension in the variable called filecount and the name of
' each file in an array called FileArray$(x).

FileTypeSearch SourceFolder$, ".CAB", "N"

' NumberOfFiles is only a temporary value returned by the FileTypeSearch subroutime. Save this to FileCount now so that NumberOfFiles
' is available to be changed by the subroutine again

FileCount = NumberOfFiles

Cls

If FileCount = 0 Then
    Color 14, 4: Print "No files with the .CAB extension were found.": Color 15
    Print "Please specify another folder."
    Pause
    GoTo ExpandDrivers
End If

' Initialize arrays

ReDim FileArray(FileCount) As String
ReDim SourceFileNameOnly(FileCount) As String

' Take the temporary array called TempArray$() and save the values in FileArray$()

For x = 1 To FileCount
    FileArray$(x) = TempArray$(x)
Next x

' We already have the names of all the CAB images to be update in the FileArray$() variables. However,
' we also want to have the name of the files with the path stripped out. We are going to store the file
' names without a path in the array called SourceFileNameOnly$().

For x = 1 To FileCount
    FileArray$(x) = TempArray$(x)
    SourceFileNameOnly$(x) = Mid$(FileArray$(x), (_InStrRev(FileArray$(x), "\") + 1))
Next x

GetDestinationPath3:

' Now that we have a valid source directory and we know that there are CAB files
' located there, ask the user for the location where we should save the expanded
' files.

DestinationFolder$ = "" ' Set initial value

Do
    Cls
    Print "Enter the destination path. Do not use the same path as the location of the .CAB files. A subfolder of that"
    Print "location is okay."
    Print
    Line Input "Enter the destination path: ", DestinationFolder$
Loop While DestinationFolder$ = ""

CleanPath DestinationFolder$
DestinationFolder$ = Temp$ + "\"

' We don't want user to specify the root of a drive

If Len(DestinationFolder$) = 3 Then
    Cls
    Color 14, 4: Print "Please do not specify the root directory of a drive.": Color 15
    Pause
    GoTo GetDestinationPath3
End If

' Verify that the path specified exists.

If Not (_DirExists(DestinationFolder$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(DestinationFolder$)) Then
        Cls
        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
        Print
        Print "Please recheck the path you have specified and try again."
        Pause
        GoTo GetDestinationPath3
    End If
End If

' If we have arrived here it means that the destination path already exists
' or we were able to create it successfully.

' The first time into the update process, we want the MainLoopCouter to be equal to 1. Then, with each loop through
' the process we will increment the counter. See the comments below for the purpose of the AutoCleanup variable.

MainLoopCount = 1 ' Set initial value

' Create the folders we need for the project.

For x = 1 To FileCount
    Cls
    Print "Processing .CAB file"; x; " of"; FileCount
    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "\" + SourceFileNameOnly$(x) + Chr$(34) + " > NUL"
    Shell Cmd$
    Cmd$ = "expand " + Chr$(34) + FileArray$(x) + Chr$(34) + " -f:*.* " + Chr$(34) + DestinationFolder$ + SourceFileNameOnly$(x) + Chr$(34) + " > NUL"
    Shell Cmd$
Next x

' If we reach point, then we have finished updating all images that need to be updated.
' Inform the user that we are done, then return to the main menu.

' Clear the keyboard buffer

_KeyClear

Cls
Print "That's all!"
Print "All .CAB files have been processed."
Print
Print "Your extracted .CAB file contents are located in the location you specified with a subfolder having the same name"
Print "as the original .CAB file, minus the .CAB extension."
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' ********************************
' * Create a Virtual Disk (VHDX) *
' ********************************

CreateVHDX:

VHDXPath$ = "" ' Set initial value

Do
    Cls
    Print "Please specify the location where you would like to create the Virtual Hard Disk. If the path does not exist, we will"
    Print "try to create it. ";: Color 0, 10: Print "Do not include a file name.": Color 15
    Print
    Line Input "Please enter path: ", VHDXPath$
Loop While VHDXPath$ = ""

' Remove trailing backslash.

CleanPath VHDXPath$
VHDXPath$ = Temp$

' Verify that the path specified exists.

If Not (_DirExists(VHDXPath$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cmd$ = "md " + Chr$(34) + VHDXPath$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(VHDXPath$)) Then
        Cls
        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
        Print
        Print "Please recheck the path you have specified and try again."
        Pause
        GoTo CreateVHDX
    End If
End If

' If we have arrived here it means that the destination path already exists
' or we were able to create it successfully.

' Get a name for the Virtual Hard Disk file.

VHDXFileName$ = "" ' Set initial value
VHDXSize = 0 ' Set initial value

Do
    Cls
    Print "Please provide a name for the Virtual Hard disk file. ";: Color 0, 10: Print "Do not include a file extension.": Color 15
    Print
    Line Input "Enter file name: ", VHDXFileName$
Loop While VHDXFileName$ = ""

Do
    Cls
    Print "Enter size of Virtual Disk ";: Color 0, 10: Print "in MB";: Color 14, 4: Print " NOT ";: Color 0, 10: Print "in GB!";: Color 15: Print ": ";
    Input "", VHDXSize
Loop While VHDXSize = 0

' We need to strip the leading space from the size, so we are going to convert the size to a string.

VHDXSizeString$ = Str$(VHDXSize)
VHDXSizeString$ = Right$(VHDXSizeString$, (Len(VHDXSizeString$) - 1))

GetVHDXLetter:

Cls
Print "What drive letter do you want to assign to the Virtual Disk. Enter only the letter, ";: Color 0, 10: Print "no colon (:)";: Color 15: Print ": ";
Input "", VHDXLetter$
VHDXLetter$ = UCase$(VHDXLetter$)

If (Len(VHDXLetter$) > 1) Or (VHDXLetter$) = "" Or ((Asc(VHDXLetter$)) < 65) Or ((Asc(VHDXLetter$)) > 90) Then
    Print
    Color 14, 4: Print "That was not a valid entry.";: Color 15: Print " Please try again."
    Print
    GoTo GetVHDXLetter
End If

If _DirExists(VHDXLetter$ + ":") Then
    Print
    Color 14, 4: Print "That drive letter is already in use.": Color 15
    Pause
    GoTo GetVHDXLetter
End If

' Create the Virtual Disk

Cls
Print "We are now creating the disk. This could possibly take a little while."
Open "TEMP.BAT" For Output As #1
Print #1, "@echo off"
Print #1, "(echo create vdisk file="; Chr$(34); VHDXPath$; "\"; VHDXFileName$; ".VHDX"; Chr$(34); " type=expandable maximum="; VHDXSizeString$
Print #1, "echo select vdisk file="; Chr$(34); VHDXPath$; "\"; VHDXFileName$; ".VHDX"; Chr$(34)
Print #1, "echo attach vdisk"
Print #1, "echo create partition primary"
Print #1, "echo format fs=ntfs quick"
Print #1, "echo assign letter="; VHDXLetter$
Print #1, "echo exit"
Print #1, ") | diskpart > NUL"
Close #1
Shell "TEMP.BAT"
If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

' Clear the keyboard buffer

_KeyClear

Cls
Print "Virtual Hard Disk has been created. The virtual disk is located here:"
Print
Color 10: Print VHDXPath$; "\"; VHDXFileName$; ".VHDX": Color 15
Print
Print "Please note that this disk was created as an expandable disk so the initial size of the file may appear much"
Print "smaller than the size that you specified."
Print
Print "If you created this disk on removable media, eject the Virtual Disk before you eject the removable media."
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' *******************************************************************
' * Create a VHD, deploy Windows to it, and add it to the boot menu *
' *******************************************************************

AddVHDtoBootMenu:

Cls
Print "This routine will create a virtual hard disk (VHD), deploy Windows to it, and add it to the boot menu of your system."
Print
Print "To use this routine, you will need to know the index number of the edition of Windows to be installed. If you do"
Print "not already know the index number of the edition that you want to install, please stop the program now, rerun it,"
Print "and select the option from the main menu called "; Chr$(34); "Get image info - display basic info for each edition in an ISO"
Print "image and display Windows build number"; Chr$(34); ". This will display each edition of Windows in the ISO image as well as the"
Print "index number associated with each edition. Once you have the index number you can rerun this routine."
Print
Print "Note: If you run the Get image info routine, when asked if you want to show information for the boot.wim and the"
Print "winre.wim, you can respond "; Chr$(34); "No"; Chr$(34); "."
Pause
Cls
Print "We need to know what ISO image contains the edition of Windows that you want to install to the VHD. Please provide"
Print "the full path, ";: Color 0, 10: Print "including the file name";: Color 15: Print ", to this ISO image."
Print
Line Input "Enter the full path to the ISO image: ", SourceImage$
CleanPath SourceImage$
SourceImage$ = Temp$

If Not (_FileExists(SourceImage$)) Then
    Cls
    Color 14, 4: Print "No such file exists.": Color 15
    Pause
    GoTo AddVHDtoBootMenu
End If

' Check image to determine the architecture
' If dual architecture, determine if x64 or x64 image is to be selected
' Set imagepath to \sources\install.wim, \x64\sources\install.wim, or \x86\sources\install.wim.

GetVHDDestination:

Do
    Cls
    Print "Enter the full path to the location where we should create the VHD ";: Color 0, 10: Print "not including a filename";: Color 15: Print "."
    Print "Please ensure that the destination is formatted with NTFS."
    Print
    Color 0, 10: Print "IMPORTANT:";: Color 15: Print " Do not install to a drive that is BitLocker encrypted other than the C: drive!"
    Print
    Line Input "Enter the path: ", Destination$
Loop While Destination$ = ""

CleanPath Destination$
Destination$ = Temp$ + "\"

' Check to see if the destination is valid

Cls
Print "Performing a check to see if the destination you specified is valid."
Print
Print "Please standby..."
DriveLetter$ = Left$(Destination$, 2)
RemovableDiskCheck DriveLetter$
DestinationIsRemovable = IsRemovable

Select Case DestinationIsRemovable
    Case 2
        Cls
        Color 14, 4: Print "This is not a valid disk.";: Color 15: Print " Please specify another location."
        Pause
        GoTo GetVHDDestination
    Case 0, 1
        ' if the returned value was a 0 or 1, no action is necessary. The program will continue normally.
End Select

' Verify that the path specified exists.

If Not (_DirExists(Destination$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cmd$ = "md " + Chr$(34) + Destination$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(Destination$)) Then
        Cls
        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
        Print
        Print "Please recheck the path you have specified and try again."
        Pause
        GoTo GetVHDDestination
    End If
End If

' If we have arrived here it means that the destination path already exists
' or we were able to create it successfully.

Do
    Cls
    Print "Enter the filename to assign to the VHD ";: Color 0, 10: Print "not including a file extension";: Color 15: Print "."
    Print
    Line Input "Enter the filename: ", VHDFilename$
Loop While VHDFilename$ = ""

' Build the full path including the filename

Destination$ = Destination$ + "\" + VHDFilename$ + ".VHDX"

' When building the path, we may end up with a situation where we get an extra backslash (\) in the path. Check for
' and correct this situation now.

If InStr(Destination$, "\\") Then
    Destination$ = Left$(Destination$, InStr(Destination$, "\\")) + (Right$(Destination$, Len(Destination$) - InStr(Destination$, "\\") - 1))
End If

If _FileExists(Destination$) Then
    Cls
    Color 14, 4: Print "Warning!";: Color 15: Print " That file already exists."
    Print
    Print "Either specify another location for the file, or specify a different filename. We will now take you back to re-enter"
    Print "the path and filename again."
    Pause
    GoTo GetVHDDestination
End If

Cls
Input "Enter the VHD size in MB: ", VHDSize

GoSub SelectAutoOrManual2:

DriveLetter$ = Letter$(1)

Cls
Print "Enter the description to be displayed in the boot menu. Example: Win 10 Pro (VHD)"
Print
Line Input "Enter description: ", Description$
Cls
Print "Please standby while we determine the architecture of the selected ISO image..."

' We now need to determine the architecture type of the selected image file. We'll simply assume index #1 since an image will always
' have at least an index 1.

DetermineArchitecture SourceImage$, 1

Select Case ImageArchitecture$
    Case "x64", "x86"
        ImagePath$ = "sources\install.wim"
    Case "DUAL"
        ChooseArchitecture:
        Cls
        Print "Your ISO image is a dual architecture image (it contains both x64 and x86 images)."
        Print
        Print "Which architecture type do you want to use? Enter "; Chr$(34);: Color 0, 10: Print "x64";: Color 15: Print Chr$(34); " or "; Chr$(34);: Color 0, 10: Print "x86";: Color 15: Print Chr$(34);: Print ": ";
        Input "", Arc$
        If ((Arc$ <> "x64") And (Arc$ <> "x86")) Then
            Cls
            Color 14, 4: Print "That is not a valid response!": Color 15
            Pause
            GoTo ChooseArchitecture
        End If
        Select Case Arc$
            Case "x64"
                ImagePath$ = "x64\sources\install.wim"
            Case "x86"
                ImagePath$ = "x86\sources\install.wim"
        End Select
    Case "NONE"
        Cls
        Color 14, 4: Print "The image specified appears to be invalid.": Color 15
        Pause
        ChDir ProgramStartDir$: GoTo BeginProgram
End Select

GetIndexforVHDDeploy:

Cls
Print "Please enter an index number. If you need to determine what index number is associated with the edition of Windows"
Print "that you want to apply to the VHD, then cancel out of this routine and run the option "; Chr$(34); "Get image info - display basic"
Print "info for each edition in an ISO image"; Chr$(34); " from the main menu."
Print
Input "Enter the index number: ", Index$

' In addition to having the index as a string, we want an integer value. We'll save that to IndexVal

IndexVal = Val(Index$)

If IndexVal = 0 Then
    Cls
    Print "Please enter a valid index number."
    Pause
    GoTo GetIndexforVHDDeploy
End If

GetVHD_Type:

Cls
Print "Enter a ";: Color 0, 10: Print " 1 ";: Color 15: Print " to create an MBR VHD or a ";: Color 0, 10: Print " 2 ";: Color 15: Print " to create a GPT VHD."
Print "Note that MBR disks are used on BIOS based systems while GPT disks are used on UEFI based systems."
Print
Input "VHD Type: ", VHD_Type
Cls

Select Case VHD_Type
    Case 1
        Cls
        Print "Standby while we create the VHD and temporarily mount it to the drive letter "; UCase$(DriveLetter$); ":"
        GoTo CreateMBR_VHD
    Case 2

        Do
            Cls
            Print "Please enter the size for the recovery environment partition in MB. It suggested to use a minimum size of 500 MB, but"
            Print "if you have sufficient space, you may want to consider a larger size such as 2048 MB (2 GB) in order to allow for"
            Print "sufficient space when performing an upgrade to a new version of Windows."
            Print
            Input "Enter size of recovery environment partition in MB: ", WinREPartitionSize
        Loop While WinREPartitionSize = 0

        Cls
        Print "Standby while we create the VHD and temporarily mount it to the drive letter "; UCase$(DriveLetter$); ":"
        GoTo CreateGPT_VHD
End Select

' If we arrive here, then a valid choice for the VHD Type was not made

Cls
Color 14, 4: Print "You did not make a valid choice for the VHD type.": Color 15
Print "Please enter a 1 for an MBR VHD or a 2 for a GPT VHD."
Pause
GoTo GetVHD_Type

CreateMBR_VHD:

Open "TEMP.BAT" For Output As #1
Print #1, "@echo off"
Print #1, "(echo create vdisk file="; Chr$(34); Destination$; Chr$(34); " maximum="; VHDSize; " "; "type=expandable"
Print #1, "echo attach vdisk"
Print #1, "echo create part primary"
Print #1, "echo format quick label="; Chr$(34); "Windows"; Chr$(34)
Print #1, "echo assign letter="; DriveLetter$
Print #1, "echo exit"
Print #1, "echo ) | diskpart > NUL"
Close #1
Shell "TEMP.BAT"
Kill "TEMP.BAT"
GoTo VHD_Created

CreateGPT_VHD:

Open "TEMP.BAT" For Output As #1
Print #1, "@echo off"
Print #1, "(echo create vdisk file="; Chr$(34); Destination$; Chr$(34); " maximum="; VHDSize; " "; "type=expandable"
Print #1, "echo attach vdisk"
Print #1, "echo convert gpt"
Print #1, "echo create partition efi size=260"
Print #1, "echo format quick fs=fat32 label="; Chr$(34); "System"; Chr$(34)
Print #1, "echo create partition msr size=128"
Print #1, "echo create partition primary"
Print #1, "echo shrink minimum="; LTrim$(Str$(WinREPartitionSize))
Print #1, "echo format quick fs=ntfs label="; Chr$(34); "Windows"; Chr$(34)
Print #1, "echo assign letter="; DriveLetter$
Print #1, "echo create partition primary"
Print #1, "echo format quick fs=ntfs label="; Chr$(34); "WinRE"; Chr$(34)
Print #1, "echo set id="; Chr$(34); "de94bba4-06d1-4d40-a16a-bfd50179d6ac"; Chr$(34)
Print #1, "echo exit"
Print #1, "echo ) | diskpart > NUL"
Close #1
Shell "TEMP.BAT"
Kill "TEMP.BAT"
GoTo VHD_Created

VHD_Created:

MountISO SourceImage$
CDROM$ = MountedImageDriveLetter$
Print "Applying the image to the VHD. You can view the progress below."
Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /apply-image /imagefile:" + Chr$(34) + CDROM$ + "\" + ImagePath$ + Chr$(34) + " /index:" + Index$ + " /applydir:" + DriveLetter$ + ":\"
Shell Chr$(34) + Cmd$ + Chr$(34)
Print "Dismounting the image"
Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourceImage$ + "'" + Chr$(34) + Chr$(34) + " > NUL"
Shell Cmd$

' If BitLocker is enabled on C:, it needs to be temporarily suspended until the next reboot
' to prevent the changes by the bcdboot and bcdedit commands from causing the system to ask
' for the recovery key.

' If BitLocker is enabled we are suspending it temporarily (until the next boot into this instace of Windows)

Cmd$ = "manage-bde -protectors -disable C: > NUL"
Shell Cmd$
Print "Updating boot information"
Cmd$ = "bcdboot " + DriveLetter$ + ":\Windows > NUL"
Shell Cmd$
Cmd$ = "bcdedit /set {default} description " + Chr$(34) + Description$ + Chr$(34) + " > NUL"
Shell Cmd$

' Clear keyboard buffer

_KeyClear

Print
Print "*************************************************"
Print "* Done. Windows has been deployed to a VHD file *"
Print "* and the host boot menu has been updated.      *"
Print "*************************************************"
Print
Print "Please note that the new instance of Windows deployed to the VHD has been made the default boot option. This can be"
Print "helpful because the system may need to be rebooted multiple times while you configure this instance of Windows. You"
Print "can change this at any time by booting your primary Windows installation, then modifying the boot options on the"
Print "boot tab of the MSCONFIG utility. Just make sure to suspend BitLocker first if it is enabled on the C: drive. It"
Print "will be automatically reenabled the next time you boot into your primary Windows installation."
Print
Print "Please note if the OS that you are currently running has BitLocker enabled on the C: drive, we have temporarily"
Print "suspended BitLocker to avoid having Windows ask you for the BitLocker key as a result of the changes that we just"
Print "made. The next time you boot into this instance of Windows (not the VHD you just created), BitLocker will be resumed."
Pause
Cls
Print "If at any time you wish to delete the newly created instance of Windows, please follow these steps:"
Print
Print "1) If BitLocker is enabled on C: of this Windows installation, suspend BitLocker either via the GUI or run this"
Print "   command:"
Print
Print "   manage-bde -protectors -disable C:"
Print
Print "   This will suspend BitLocker until the system is rebooted into this installation of Windows."
Print
Print "2) Run MSCONFIG, go to the boot tab, set your primary Windows installation to the default, and delete the entry for the"
Print "   installation of Windows deployed to the VHD."
Print
Print "   Note that even when you suspend BitLocker, Windows will warn you that you will be asked to supply the recovery key"
Print "   when making the changes in MSCONFIG. If you suspended BitLocker, you can ignore this message."
Print
Print "3) You can now delete the VHD to which you deployed Windows."
Pause
Cls
Print "If you want to connect the VHD to a Hyper-V VM, follow these steps:"
Print
Print "1) Create a new VM but do not install Windows to it."
Print
Print "2) When creating the VM, don't create a new virtual disk. Instead, tell it to use an existing virtual disk and point"
Print "   it to the virtual disk file that you created."
Print
Print "3) Boot the new VM from Windows install media (an ISO image)."
Print
Print "4) Open a command prompt by pressing SHIFT + F10 and then run "; Chr$(34); "diskpart"; Chr$(34); "."
Print
Print "5) Within diskpart, run "; Chr$(34); "list vol"; Chr$(34); " and note the drive letter for the Windows partition (for example, C:)."
Print
Print "6) Exit from diskpart using the "; Chr$(34); "exit"; Chr$(34); " command."
Print
Print "7) Run "; Chr$(34); "C:\Windows\System32\bcdboot C:\Windows"; Chr$(34); " - Replace C: with the drive letter that "; Chr$(34); "list vol"; Chr$(34); " showed."
Pause

ChDir ProgramStartDir$: GoTo BeginProgram

' Local Subroutine

' Get drive letter to assign to each partition
' The user can choose to manually assign drive letters to the partitions being created on the bootable media
' or allow the program to automatically assign drive letters.

SelectAutoOrManual2:

TotalPartitions = 1 ' For this routine, we need to get one drive letter

ReDim Letter(TotalPartitions) As String

Cls
Print "The program will automatically assign a temporary drive letter to the VHD. However, if you prefer, you can manually"
Print "assign a drive letter."
Print
Input "Do you want to manually assign a drive letter"; ManualAssignment$

YesOrNo ManualAssignment$

Select Case YN$
    Case "Y"
        GoTo ManualAssign2
    Case "N"
        GoTo AutoAssign2
    Case "X"
        GoTo SelectAutoOrManual2
End Select

ManualAssign2:

' Allow the user to manually choose drive letter

For x = 1 To TotalPartitions
    Do

        RepeatLetter2:

        Cls
        Print "Enter the drive letter to assign. Enter only the letter without a colon (:)."
        Print
        Letter$(x) = "" ' Set initial value
        Input "Enter the drive letter to use: ", Letter$(x)
    Loop While Letter$(x) = ""

    Letter$(x) = UCase$(Letter$(x))

    If (Len(Letter$(x)) > 1) Or (Letter$(x)) = "" Or ((Asc(Letter$(x))) < 65) Or ((Asc(Letter$(x))) > 90) Then
        Print
        Color 14, 4: Print "That was not a valid entry. Please try again.": Color 15
        Print
        GoTo RepeatLetter2
    End If

    If _DirExists(Letter$(x) + ":") Then
        Print
        Color 14, 4: Print "That drive letter is already in use.": Color 15
        Pause
        GoTo RepeatLetter2
    End If

Next x

GoTo LetterAssignmentDone2

AutoAssign2:

' Automatically assign drive letters

' To auto assign drive letters, we go through a loop checking to see if drive letters C:, D:, E:, etc. are already in use.
' If in use, we move on to the next drive letter.

' Before we search for available drive letters, we will call a subroutine that looks for removable media with volumes that have
' a status of "Unusable". This happens when a removable media drive such as a USB Flash Drive (UFD) has had a "clean" operation
' performed on it in Diskpart. The drive now has now partitions but it still shows up in File Explorer with a drive letter. This
' can cause difficulties for drive letter detection.

Print
Print "Looking for available drive letters."

CleanVols

LettersAssigned = 0 ' Keep track of how many drive letters were assigned. Once equal 1, we are done.

Restore DriveLetterData

For y = 1 To 24
    Read Letter$(LettersAssigned + 1)
    Cmd$ = "dir " + Letter$(LettersAssigned + 1) + ":\ > DriveStatus.txt 2>&1"
    Shell Cmd$
    ff = FreeFile
    Open "DriveStatus.txt" For Input As #ff
    FileLength = LOF(ff)
    Temp$ = Input$(FileLength, ff)
    Close #ff
    Kill "DriveStatus.txt"

    If (InStr(Temp$, "The system cannot find the path specified")) Then
        LettersAssigned = LettersAssigned + 1
    End If

    If LettersAssigned = TotalPartitions Then GoTo LetterAssignmentDone2

Next y

' The FOR loop should only complete if we run out of drive letters. We need to warn the user about this and how they can correct
' the issue. The program will then end.

Cls
Print "No drive letter was available to assign!"
Print
Print "Solution: Please free up a drive letter and then re-run this program."
Print "The program will now end."
Pause
System

LetterAssignmentDone2:

Return


' *******************************************************************
' * Create a generic ISO image and inject files and folders into it *
' *******************************************************************

CreateISOImage:

' Initialize variables

SourcePath$ = ""
VolumeName$ = ""

Do
    Cls

    ' Get the location to the files / folders that should be injected into an ISO file.

    Line Input "Enter the path containing the data to place into an ISO image: ", SourcePath$
Loop While SourcePath$ = ""

CleanPath SourcePath$
SourcePath$ = Temp$

' Verify that the path specified exists.

If Not (_DirExists(SourcePath$)) Then
    Cls
    Color 14, 4: Print "The location that you specified does not exist or is not valid.": Color 15
    Pause
    GoTo CreateISOImage
End If

getdestination:

DestinationPath$ = "" ' Set initial value

Do
    Cls
    Line Input "Enter the destination path. This is the path only without a file name: ", DestinationPath$
Loop While DestinationPath$ = ""

CleanPath DestinationPath$
DestinationPath$ = Temp$ + "\"

' Verify that the path specified exists.

If Not (_DirExists(DestinationPath$)) Then

    ' The destination path does not exist. We will now attempt to create it.

    Cls

    Print "Destination path does not exist. Attempting to create it..."
    Cmd$ = "md " + Chr$(34) + DestinationPath$ + Chr$(34)
    Shell _Hide Cmd$

    ' Checking for existance of folder again again to see if we were able to create it.

    If Not (_DirExists(DestinationPath$)) Then
        Print
        Color 14, 4: Print "We were not able to create the destination folder.": Color 15
        Print "Please recheck the path you have specified and try again."
        Pause
        GoTo getdestination
    End If
End If

' If we have arrived here it means that the destination path already exists
' or we were able to create it successfully.

DestinationFileName$ = "" ' Set initial value

' Get the name of the ISO image that we are creating.

Do
    Cls
    Print "Enter the name of the file to create, ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension: ";: Line Input "", DestinationFileName$
Loop While DestinationFileName$ = ""

DestinationPathAndFile$ = DestinationPath$ + DestinationFileName$ + ".iso"

' Get the volume name for the ISO image

GetVolumeName2:

Cls
Input "Enter the volume name to give the ISO image or press Enter for none: ", VolumeName$

If Len(VolumeName$) > 32 Then
    Print
    Color 14, 4: Print "That volume name is invalid!";: Color 15: Print " The volume name is limited to 32 characters."
    Pause
    GoTo GetVolumeName2
End If

' Build the command that needs to be run to create the ISO image.

Do
    _Limit 10
    CurrentTime$ = Date$ + "," + Left$(Time$, 5)
    Select Case Right$(CurrentTime$, 8)
        Case "23:59:58", "23:59:59"
            Midnight = 1
        Case Else
            Midnight = 0
    End Select
Loop While Midnight = 1

Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -t" + CurrentTime$ + " -o -m -h -u2 -udfver102 -l" + CHR$(34) + VolumeName$ + CHR$(34) + " " + CHR$(34) + SourcePath$_
+ CHR$(34) + " " + CHR$(34) + DestinationPathAndFile$ + CHR$(34) + " > NUL 2>&1"

' Create the ISO image

Cls
Print "Creating the image. Please standby..."
Shell Chr$(34) + Cmd$ + Chr$(34)
Print
Print "ISO Image created."

' Clear the keyboard buffer

_KeyClear

Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' *****************************
' * Cleanup files and folders *
' *****************************

GetFolderToClean:

DestinationPath$ = "" ' Set initial value

Do
    Cls
    Line Input "Please enter the full path to the project folder to be cleaned: ", DestinationPath$
Loop While DestinationPath$ = ""

CleanPath DestinationPath$
DestinationPath$ = Temp$ + "\"

' We don't want user to specify the root of a drive

If Len(DestinationFolder$) = 3 Then
    Cls
    Color 14, 4: Print "Please do not specify the root directory of a drive.": Color 15
    Pause
    GoTo GetFolderToClean
End If

Cleanup DestinationPath$

' Clear the keyboard buffer

_KeyClear

If CleanupSuccess = 1 Then
    Cls
    Print "The contents of the folder ";: Color 10: Print DestinationPath$;: Color 15: Print " were ";: Color 0, 10: Print "successfully";: Color 15: Print " cleaned."
    Print
    Print "If any log files were present, we have left them alone and they will still be available."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

Cls
Color 14, 4: Print "There was a problem performing the cleanup.";: Color 15: Print " You may need to perform a manual cleanup."
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' ************************************
' * Unattended Answer File Generator *
' ************************************

AnswerFileGen:

' Introduction.

Cls
Color 0, 10: Print "Welcome to the Answer File Generator!": Color 15
Print
Print "Welcome to the Answer File Generator. On the screens that follow, we will ask for the information needed to generate"
Print "an answer file (autounattend.xml) for your system to automate the installation of Windows."
Print
Print "When we ask a question that needs a yes or no response, you can respond with a single letter (Y or N) or the full"
Print "word (YES or NO). Case does not matter."
Print
Print "You will be asked what size you wish to make various partitions. When entering a size, always specify the size in MB."
Print "For example, to specify a size of 100 GB you would enter 100000 since 100 GB equals 100000 MB. Enter only a number"
Print "for these responses."
Print
Print "Please note that at this time, this program is designed to work only with US English language editions of Windows."
Pause

' Start gathering the information needed to generate an answer file.

' Ask for system type. Loop until a valid response is received.

Do
    Cls
    Color 0, 10: Print "System Type": Color 15
    Print
    Print "We need to know what type of system this answer file is for, either a BIOS based system or a UEFI based system."
    Print "Each system type has different requirements that we need to account for."
    Print
    Print "Enter a ";: Color 10: Print "1 for UEFI";: Color 15: Print " or a ";: Color 10: Print "2 for BIOS";: Color 15: Print ": ";
    Input "", SystemType$

    Select Case SystemType$
        Case "1"
            SystemType$ = "UEFI"
            GoTo GetInfoForUefiSys
        Case "2"
            SystemType$ = "BIOS"
            GoTo CommonInfo
        Case Else
            Print
            Color 14, 4: Print "INVALID ENTRY";: Color 15: Print ": Please enter only the number 1 or 2."
            Pause
    End Select
Loop

GetInfoForUefiSys:

' Set a default value of 260 MB for the EFI partition size.

Temp$ = "260"

Cls
Color 0, 10: Print "EFI Partition Size": Color 15
Print
Print "What size do you want to make the EFI partition? 100 MB is fine for most systems but for Advanced Format (AF)"
Print "drives this value must be at least 260 MB. To guarantee compatibility with all drives, 260 MB is recommended."
Print
Input "Enter the size of the EFI partition press ENTER for 260 MB: ", EfiParSize$

If EfiParSize$ = "" Then
    EfiParSize$ = Temp$
End If

' Set a default value of 16 MB for the MSR partition.

Temp$ = "16"

Cls
Color 0, 10: Print "Microsoft Reserved (MSR) Partition Size": Color 15
Print
Print "What size do you want to make the Mirosoft Reserved (MSR) partition? 16 MB is recommended."
Print
Input "Enter the size of the Microsoft Reserved partition (MSR) or press ENTER to use 16 MB: ", MsrParSize$

If MsrParSize$ = "" Then
    MsrParSize$ = Temp$
End If

' Ask if automatic device encryption should be bypassed. Loop until a valid response is received.

Do
    Cls
    Color 0, 10: Print "Bypass Automatic Device Encryption?": Color 15
    Print
    Print "On systems meeting certain requirements, setup will automatically initiate device encryption but you can prevent"
    Print "this from happening."
    Print
    Input "Do you want to bypass automatic device encryption"; BypassDeviceEncryption$

    YesOrNo BypassDeviceEncryption$
    BypassDeviceEncryption$ = YN$

    Select Case BypassDeviceEncryption$
        Case "Y", "N"
            Exit Do
        Case "X"
            Print
            Color 14, 4: Print "INVALID ENTRY";: Color 15: Print ": We were expecting a YES or NO response."
            Print "Reminder: Valid responses are the single letter Y or N or the words YES or NO. Case does not matter."
            Pause
    End Select
Loop

' Ask if user wants to limit the size of the Windows partition. Loop until valid response is received.

Do
    Cls
    Color 0, 10: Print "Limit the Size of the Windows Partition?": Color 15
    Print
    Print "Normally, we create the Windows partition with all space not already assigned to other partitions minus the"
    Print "size of the Windows Recovery Environment (WinRE) partition. However, there may be times where you want to limit"
    Print "the size of the Windows partition. This will leave room for pther partitions to be created. Note that if you do"
    Print "choose to limit the Windows partition size, we will create another partition using the free space. If you don't"
    Print "like this arrangement, you can easily remove that partition and replace it with as many partitions as you wish"
    Print "after Windows setup has finished."
    Print
    Input "Do you want to limit the size of the Windows partition"; LimitWinParSize$

    YesOrNo LimitWinParSize$
    LimitWinParSize$ = YN$

    Select Case LimitWinParSize$
        Case "Y"
            GoTo LimitSize
        Case "N"
            GoTo CommonInfo
        Case "X"
            Print
            Color 14, 4: Print "INVALID ENTRY";: Color 15: Print ": We were expecting a YES or NO response."
            Print "Reminder: Valid responses are the single letter Y or N or the words YES or NO. Case does not matter."
            Pause
    End Select
Loop

LimitSize:

' Ask for the size to make the Windows partition.

Cls
Color 0, 10: Print "What Size Windows Partition Would You Like?": Color 15
Print
Print "You have selected the option to limit the size of the Windows partition. What size would you like to make it?"
Print
Input "Enter the size to make the Windows partition: ", WinParSize$

' Gather info common to both BIOS and UEFI systems

CommonInfo:

' Ask for size to make the WinRE partition.

GetWinReParSize:

' Set a default size of 2 GB (2048 MB) for the WinRE partition.

Temp$ = "2048"

Cls
Color 0, 10: Print "What Size Recovery Environment Partition Would You Like?": Color 15
Print
Print "The WinRE partition should be created with a minimum of 750 MB although many people prefer to use 1000 MB since"
Print "Microsoft has been slowly increasing the amount of space used in this partition. Personally, I use 2048 MB on any"
Print "system where I am not short on space."
Print
Input "Enter the size to make the WinRE partition or press enter for 2 GB (2048 MB): ", WinReParSize$

If WinReParSize$ = "" Then
    WinReParSize$ = Temp$
End If

GetDiskId:

' Get the ID of the disk to which Windows should be installed.

Cls
Color 0, 10: Print "Disk ID for Drive Where Windows Should be Installed": Color 15
Print
Print "Setup needs to know the ID of the disk onto which it should be installed."

Color 14, 4: Print "IMPORTANT:";: Color 15: Print " The disk number that you specify here will be ";: Color 14, 4: Print "ERASED";: Color 15: Print " when Windows is installed. Do not use your running"
Print "Windows installation to try to determine the disk ID because disk IDs during Windows setup may be different than while"
Print "running Windows. If you have not already done so, you should follow these steps to determine the correct disk ID:"
Print
Print "1) Create the Windows installation media that you will use to install Windows now. Do not include an autounattend.xml"
Print "   answer file!"
Print
Print "2) Boot from that media."
Print
Print "3) At the very first static screen, press SHIFT + F10 to open a command prompt."
Print
Print "4) At the command prompt, run ";: Color 10: Print "diskpart";: Color 15: Print "."
Print
Print "5) Once diskpart has started, run the command ";: Color 10: Print "list disk";: Color 15: Print ". Note the disk ID (disk number) of the disk to which you"
Print "   will install Windows. If the information shown is not enough to allow you to determine the correct disk, then"
Print "   select a disk and show details for that disk to get more info. You can do this for as many disks as needed."
Print "   EXAMPLE: ";: Color 10: Print "select disk 0";: Color 15: Print ", ";: Color 10: Print "detail disk";: Color 15: Print "."
Print "6) Run ";: Color 10: Print "exit";: Color 15: Print " twice to close diskpart and the command prompt."
Print
Print "7) Reboot the system back into Windows."
Print
Color 14, 4: Print "IMPORTANT:";: Color 15: Print " Once the disk ID is determined, don't add or remove drives as the disk ID may then change!"
Pause

Cls
Color 0, 10: Print "Disk ID for Drive Where Windows Should be Installed": Color 15
Print
Input "Enter the disk ID of the disk to which Windows should be installed: ", DiskIdTarget$

' Ask if Windows 11 system requirements should be bypassed.

BypassRequirements:

Do
    Cls
    Color 0, 10: Print "Bypass Windows 11 System Requirements?": Color 15
    Print
    Print "Windows 11 has specific requirements such as the presence of a TPM and the availability of Secure Boot. However,"
    Print "many, if not most systems that do not meet these requirements can run Windows 11 just fine. By choosing this"
    Print "option, we bypass these checks performed by Windows setup. Note that with thjis option enabled, you can still use"
    Print "this answer file on systems that meet the requirements or even with Windows 10 installation media."
    Print
    Input "Do you want to bypass the Windows 11 system requirements check"; BypassWinRequirements$

    YesOrNo BypassWinRequirements$
    BypassWinRequirements$ = YN$

    Select Case BypassWinRequirements$
        Case "Y", "N"
            GoTo QualityUpdates
        Case "X"
            Print
            Color 14, 4: Print "INVALID ENTRY";: Color 15: Print ": We were expecting a YES or NO response."
            Print "Reminder: Valid responses are the single letter Y or N or the words YES or NO. Case does not matter."
            Pause
    End Select
Loop

' Ask if the downloading and installation of Windows quality updates during setup should be bypassed.

QualityUpdates:

Do
    Cls
    Color 0, 10: Print "Prevent Quality Updates During Setup?": Color 15
    Print
    Print "Windows can install quality updates during installation. This can take a considerable amount of time and it can also"
    Print "result in you ending up with a different build of Windows than intended."
    Print
    Input "Do you want to prevent checks for quality updates during Windows installation"; BypassQualityUpdatesDuringOobe$

    YesOrNo BypassQualityUpdatesDuringOobe$
    BypassQualityUpdatesDuringOobe$ = YN$

    Select Case BypassQualityUpdatesDuringOobe$
        Case "Y", "N"
            GoTo GetProdKey
        Case "X"
            Print
            Color 14, 4: Print "INVALID ENTRY";: Color 15: Print ": We were expecting a YES or NO response."
            Print "Reminder: Valid responses are the single letter Y or N or the words YES or NO. Case does not matter."
            Pause
    End Select
Loop

GetProdKey:

' Default to using the Windows Pro edition key.

Temp$ = "VK7JG-NPHTM-C97JM-9MPGT-3V66T"

Do
    Cls
    Color 0, 10: Print "Provide Qindows Generic Installation Key": Color 15
    Print
    Print "We need the generic product key for the edition of Windows that you wish to install. Below are the most commonly"
    Print "used product keys. For other keys, please visit this link:"
    Print
    Print "https://www.elevenforum.com/t/generic-product-keys-to-install-or-upgrade-windows-11-editions.3713/"
    Print
    Print "Windows 10 or 11 Home Single Language:  BT79Q-G7N6G-PGBYW-4YWX6-6F4BT"
    Print "Windows 10 or 11 Home:                  YTMG3-N6DKC-DKB77-7M9GH-8HVX7"
    Print "Windows 10 or 11 Pro:                   VK7JG-NPHTM-C97JM-9MPGT-3V66T"
    Print

    Input "Provide the key for the edition being installed or press ENTER to use the key for Pro: ", ProductKey$

    If ProductKey$ = "" Then
        ProductKey$ = Temp$
    End If

    If Len(ProductKey$) = 29 Then Exit Do

    Print
    Color 14, 4: Print "INVALID ENTRY";: Color 15: Print ": The key is 29 characters long (including dashes)."
    Print "Reminder: Please try again."
    Pause
Loop

' Get UserLocale.

Do
    Cls
    Color 0, 10: Print "Provide UserLocale": Color 15
    Print
    Print "Normally you will provide a UserLocale of ";: Color 10: Print "en-US";: Color 15: Print "but you can also specify ";: Color 10: Print "en-001";: Color 15: Print "."
    Print "Using a UserLocale of ";: Color 10: Print "en=001";: Color 15: Print "sets the region to ";: Color 10: Print "WORLD";: Color 15: Print ". This has the effect"
    Print "causing the Start screen to be very sparsely populated. However, if you choose this option there will be a couple"
    Print "of quick operations thatyou need to perform. We'll tell you what you need to do after the answer file is created."
    Print
    Print "Enter a ";: Color 10: Print "1 for en-US";: Color 15: Print " or ";: Color 10: Print "2 for en-001";: Color 15: Print ": ";
    Input "", Temp$

    Select Case Temp$
        Case "1"
            UserLocale$ = "en-US"
            GoTo GetName
        Case "2"
            UserLocale$ = "en-001"
            GoTo GetName
    End Select

Loop

' Get user name

GetName:

Cls
Color 0, 10: Print "Provide UserName": Color 15
Print
Print "A local user account will be created and added to the Administrators group. We need the User Name to be"
Print "used for this."
Print
Print "On the next screen, you will be asked for the Full Name or Display Name. This is the full name or friendly"
Print "name that is displayed in places such as the lock screen."
Print
Print "Example: You might have a user name of ";: Color 10: Print "WinUser";: Color 15: Print " and a full name of ";: Color 10: Print "Windows User";: Color 15: Print "."
Print
Input "Enter user name: ", UserName$

' Get full name / display name.

GetDisplayName:

Cls
Color 0, 10: Print "Provide Display Name / Full Name": Color 15
Print
Print "This is the full name associated with the user name that you just created as described on the previous screen."
Print
Input "Enter the display name: ", DisplayName$

' Get time zone

GetTimezone:

' Set a defult of Central Standard Time

Temp$ = "Central Standard Time"

Cls
Color 0, 10: Print "Provide the Time Zone Where System Will be Located": Color 15
Print
Print "Please enter the time zone that this computer will be located in. For example, ";: Color 10: Print "Central Standard Time";: Color 15: Print ". To get a list of"
Print "valid time zones, run the command ";: Color 10: Print "tzutil /L";: Color 15: Print ". The second line of each group is the name that you can specify here."
Print
Print "Enter time zone or press ENTER to use ";: Color 10: Print "Central Standard Time";: Color 15: Print ":";: Input " ", TimeZone$

If TimeZone$ = "" Then
    TimeZone$ = Temp$
End If

' Get Computer Name

GetComputerName:

' Set defauly computer name to an empty string.

Cls
Color 0, 10: Print "Provide the Computer Name": Color 15
Print
Print "You can provide a name to be supplied to the computer here. Note that the recommended course of action here is to"
Print "leave the computer name blank (just hit ENTER). This will cause Windows setup to generate a rando name that you can"
Print "change after installation. This allows use of the same answer file with different machines without the danger of"
Print "duplicating names."
Print
Input "Enter computer name or press ENTER for a random name: ", ComputerName$

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' All needed info has been gathered. Begin generation of the answer file '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Save the answer file to the location where this program was run from.

AnswerFilePath$ = ProgramStartDir$ + "autounattend.xml"

ff = FreeFile
Open AnswerFilePath$ For Output As #ff

Cls
Print "Generating answer file. Please standby..."

' The following section is included for both BIOS and UEFI systems.

Print #ff, "<?xml version="; Chr$(34); "1.0"; Chr$(34); " encoding="; Chr$(34); "utf-8"; Chr$(34); "?>"
Print #ff, "<unattend xmlns="; Chr$(34); "urn:schemas-microsoft-com:unattend"; Chr$(34); ">"
Print #ff, "   <settings pass="; Chr$(34); "windowsPE"; Chr$(34); ">"
Print #ff, "       <component name="; Chr$(34); "Microsoft-Windows-International-Core-WinPE"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); ">"
Print #ff, "           <SetupUILanguage>"
Print #ff, "               <UILanguage>en-US</UILanguage>"
Print #ff, "           </SetupUILanguage>"
Print #ff, "           <InputLocale>en-US</InputLocale>"
Print #ff, "           <SystemLocale>en-US</SystemLocale>"
Print #ff, "           <UILanguage>en-US</UILanguage>"
Print #ff, "           <UserLocale>"; UserLocale$; "</UserLocale>"
Print #ff, "       </component>"
Print #ff, "       <component name="; Chr$(34); "Microsoft-Windows-Setup"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); ">"
Print #ff, "           <ImageInstall>"
Print #ff, "               <OSImage>"
Print #ff, "                   <InstallTo>"
Print #ff, "                       <DiskID>"; DiskIdTarget$; "</DiskID>"

' If Windows is being installed on a UEFI system, then set the partition to which Windows should be installed to 3. On
' a BIOS based system, use partition 2.

Select Case SystemType$
    Case "UEFI"
        InstallPar$ = "3"
    Case "BIOS"
        InstallPar$ = "2"
End Select

Print #ff, "                       <PartitionID>"; InstallPar$; "</PartitionID>"
Print #ff, "                   </InstallTo>"
Print #ff, "                   <Compact>false</Compact>"
Print #ff, "               </OSImage>"
Print #ff, "           </ImageInstall>"
Print #ff, "           <UserData>"

' Add product key

Print #ff, "               <ProductKey>"
Print #ff, "                   <Key>"; ProductKey$; "</Key>"
Print #ff, "               </ProductKey>"
Print #ff, "               <AcceptEula>true</AcceptEula>"
Print #ff, "           </UserData>"

' Create a RunSynchronous block from which we can run commands
' Note that commands need to have their "Order" numbers set sequentially without skipping numbers. In order to
' facilitate this now and for possible future additions, we are setting a counter to "1" and incrementing the value
' each time a new command is added.

Phase1Commandcounter = 1

Print #ff, "           <RunSynchronous>"

If BypassWinRequirements$ = "N" GoTo BypassRequirementsDone

' These commands bypass Windows 11 system requirements

Print #ff, "               <RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <Order>"; LTrim$(Str$(Phase1Commandcounter)); "</Order>"
Print #ff, "                   <Path>reg add HKLM\System\Setup\LabConfig /v BypassTPMCheck /t reg_dword /d 0x00000001 /f</Path>"
Print #ff, "               </RunSynchronousCommand>"

Phase1Commandcounter = Phase1Commandcounter + 1

Print #ff, "               <RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <Order>"; LTrim$(Str$(Phase1Commandcounter)); "</Order>"
Print #ff, "                   <Path>reg add HKLM\System\Setup\LabConfig /v BypassSecureBootCheck /t reg_dword /d 0x00000001 /f</Path>"
Print #ff, "               </RunSynchronousCommand>"

Phase1Commandcounter = Phase1Commandcounter + 1

Print #ff, "               <RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <Order>"; LTrim$(Str$(Phase1Commandcounter)); "</Order>"
Print #ff, "                   <Path>reg add HKLM\System\Setup\LabConfig /v BypassRAMCheck /t reg_dword /d 0x00000001 /f</Path>"
Print #ff, "               </RunSynchronousCommand>"

Phase1Commandcounter = Phase1Commandcounter + 1

BypassRequirementsDone:

' Do these steps for UEFI systems

' This command performs the disk setup operations for UEFI systems only if the user wants to limit the size of the
' Windows partition. We skip this for legacy BIOS based systems.

If SystemType$ = "BIOS" GoTo BiosPartitioning
If LimitWinParSize$ = "N" GoTo CreateFullSizeWinPar

Print #ff, "               <RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <Order>"; LTrim$(Str$(Phase1Commandcounter)); "</Order>"
Print #ff, "                   <Path>cmd /c (for %a in ("; Chr$(34); "sel dis "; DiskIdTarget$; Chr$(34); " "; Chr$(34); "cle"; Chr$(34); " "; Chr$(34); "con gpt"; Chr$(34); " "; Chr$(34); "cre par efi size="; EfiParSize$; Chr$(34); " "; Chr$(34); "for quick fs=fat32"; Chr$(34); " "; Chr$(34); "cre par msr size="; MsrParSize$; Chr$(34); " "; Chr$(34); "cre par pri size="; WinParSize$; Chr$(34); " "; Chr$(34); "format quick fs=ntfs label="; Chr$(34); "Windows"; Chr$(34); ""; Chr$(34); " "; Chr$(34); "cre par pri size="; WinReParSize$; Chr$(34); " "; Chr$(34); "for quick fs=ntfs"; Chr$(34); " "; Chr$(34); "set id=de94bba4-06d1-4d40-a16a-bfd50179d6ac"; Chr$(34); " "; Chr$(34); "gpt attributes=0x8000000000000001"; Chr$(34); " "; Chr$(34); "create partition primary"; Chr$(34); " "; Chr$(34); "format quick fs=ntfs"; Chr$(34); ") do @echo %~a) &gt; X:\UEFI.txt &amp; diskpart /s X:\UEFI.txt</Path>"
Print #ff, "               </RunSynchronousCommand>"

Phase1Commandcounter = Phase1Commandcounter + 1

' This is the last command to be run so we can now close the RunSynchronous block.

Print #ff, "           </RunSynchronous>"

GoTo DonePartitioning

CreateFullSizeWinPar:

' Create a full size Windows partition

Print #ff, "               <RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <Order>"; LTrim$(Str$(Phase1Commandcounter)); "</Order>"
Print #ff, "                   <Path>cmd /c (for %a in ("; Chr$(34); "sel dis "; DiskIdTarget$; Chr$(34); " "; Chr$(34); "cle"; Chr$(34); " "; Chr$(34); "con gpt"; Chr$(34); " "; Chr$(34); "cre par efi size="; EfiParSize$; Chr$(34); " "; Chr$(34); "for quick fs=fat32"; Chr$(34); " "; Chr$(34); "cre par msr size="; MsrParSize$; Chr$(34); " "; Chr$(34); "cre par pri"; Chr$(34); " "; Chr$(34); "shr minimum="; WinReParSize$; Chr$(34); " "; Chr$(34); "for quick fs=ntfs label="; Chr$(34); "Windows"; Chr$(34); ""; Chr$(34); " "; Chr$(34); "cre par pri"; Chr$(34); " "; Chr$(34); "for quick fs=ntfs"; Chr$(34); " "; Chr$(34); "set id=de94bba4-06d1-4d40-a16a-bfd50179d6ac"; Chr$(34); " "; Chr$(34); "gpt attributes=0x8000000000000001"; Chr$(34); ") do @echo %~a) &gt; X:\UEFI.txt &amp; diskpart /s X:\UEFI.txt</Path>"
Print #ff, "               </RunSynchronousCommand>"

Phase1Commandcounter = Phase1Commandcounter + 1

' This is the last command to be run so we can now close the RunSynchronous block.

Print #ff, "           </RunSynchronous>"

GoTo DonePartitioning

BiosPartitioning:

' This was the last command to be run so we can close out the RunSynchronous block.

Print #ff, "           </RunSynchronous>"

' These are the operations needed to setup the drive for a BIOS based system.

Print #ff, "           <DiskConfiguration>"
Print #ff, "               <Disk wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <CreatePartitions>"
Print #ff, "                       <CreatePartition wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                           <Order>1</Order>"
Print #ff, "                           <Size>"; WinReParSize$; "</Size>"
Print #ff, "                           <Type>Primary</Type>"
Print #ff, "                       </CreatePartition>"
Print #ff, "                       <CreatePartition wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                           <Extend>true</Extend>"
Print #ff, "                           <Order>2</Order>"
Print #ff, "                           <Type>Primary</Type>"
Print #ff, "                       </CreatePartition>"
Print #ff, "                   </CreatePartitions>"
Print #ff, "                   <ModifyPartitions>"
Print #ff, "                       <ModifyPartition wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                           <Active>true</Active>"
Print #ff, "                           <Format>NTFS</Format>"
Print #ff, "                           <Order>1</Order>"
Print #ff, "                           <PartitionID>1</PartitionID>"
Print #ff, "                       </ModifyPartition>"
Print #ff, "                       <ModifyPartition wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                           <Format>NTFS</Format>"
Print #ff, "                           <Label>Windows</Label>"
Print #ff, "                           <Letter>C</Letter>"
Print #ff, "                           <Order>2</Order>"
Print #ff, "                           <PartitionID>2</PartitionID>"
Print #ff, "                       </ModifyPartition>"
Print #ff, "                   </ModifyPartitions>"
Print #ff, "                   <DiskID>0</DiskID>"
Print #ff, "                   <WillWipeDisk>true</WillWipeDisk>"
Print #ff, "               </Disk>"
Print #ff, "           </DiskConfiguration>"

DonePartitioning:

Print #ff, "       </component>"
Print #ff, "   </settings>"
Print #ff, "   <settings pass="; Chr$(34); "oobeSystem"; Chr$(34); ">"
Print #ff, "       <component name="; Chr$(34); "Microsoft-Windows-International-Core"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); ">"
Print #ff, "           <InputLocale>en-US</InputLocale>"
Print #ff, "           <SystemLocale>en-US</SystemLocale>"
Print #ff, "           <UILanguage>en-US</UILanguage>"

' Set user locale to en-US or en-001 depending upon what the user selected.

Print #ff, "           <UserLocale>"; UserLocale$; "</UserLocale>"

Print #ff, "       </component>"
Print #ff, "       <component name="; Chr$(34); "Microsoft-Windows-Shell-Setup"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); ">"
Print #ff, "           <OOBE>"
Print #ff, "               <HideEULAPage>true</HideEULAPage>"
Print #ff, "               <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>"
Print #ff, "               <HideOnlineAccountScreens>true</HideOnlineAccountScreens>"
Print #ff, "               <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>"
Print #ff, "               <ProtectYourPC>1</ProtectYourPC>"
Print #ff, "               <UnattendEnableRetailDemo>false</UnattendEnableRetailDemo>"
Print #ff, "           </OOBE>"
Print #ff, "           <UserAccounts>"
Print #ff, "               <LocalAccounts>"
Print #ff, "                   <LocalAccount wcm:action="; Chr$(34); "add"; Chr$(34); ">"

' We will assign the user an initial password of "Password1". The user should change this when installation completes.

Print #ff, "                       <Password>"
Print #ff, "                           <Value>Password1</Value>"
Print #ff, "                           <PlainText>true</PlainText>"
Print #ff, "                       </Password>"

' Make the local user account being created a member of the Administrators group.

Print #ff, "                       <DisplayName>"; DisplayName$; "</DisplayName>"
Print #ff, "                       <Group>Administrators</Group>"
Print #ff, "                       <Name>"; UserName$; "</Name>"
Print #ff, "                   </LocalAccount>"
Print #ff, "               </LocalAccounts>"
Print #ff, "           </UserAccounts>"

' Set time zone

Print #ff, "           <TimeZone>"; TimeZone$; "</TimeZone>"

' Add a registry entry to resolve a bug related to autologon. This answer file will autologon just one time in order
' to complete Windows setup. Later in this answer file you will see where we specify a logon count of one. The bug is
' that Windows will autologon one time more than specified. So, you would think that you could specify zero and that
' this would result in one logon. Unfortunately, the system does properly understand that zero means nerver logon.
' The registry entry works around this bug."

' Start by creating a counter called FirstLogonCommandCounter to keep count of the number of commands we are adding so
' that we assign the proper "order" to each commands. The value for "order" must increment by one each time and cannot
' skip numbers.

FirstLogonCommandCounter = 1

Print #ff, "           <FirstLogonCommands>"
Print #ff, "               <SynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <CommandLine>reg add &quot;HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon&quot; /v AutoLogonCount /t REG_DWORD /d 0 /f</CommandLine>"
Print #ff, "                   <Order>"; LTrim$(Str$(FirstLogonCommandCounter)); "</Order>"
Print #ff, "               </SynchronousCommand>"

FirstLogonCommandCounter = FirstLogonCommandCounter + 1

' Add the command below only if bypassing quality updates during Windows setup.

If BypassQualityUpdatesDuringOobe$ = "N" GoTo DoneWithQualityUpdates

Print #ff, "               <SynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <Order>"; LTrim$(Str$(FirstLogonCommandCounter)); "</Order>"
Print #ff, "                   <CommandLine>powershell.exe -Command &quot;Get-NetAdapter | ForEach-Object { Enable-NetAdapter -Name $_.Name -Confirm:$false }&quot;</CommandLine>"
Print #ff, "               </SynchronousCommand> "

FirstLogonCommandCounter = FirstLogonCommandCounter + 1

DoneWithQualityUpdates:

' Setup will not fully complete until the user logs on for the first time. We are setting a one-time automatic logon so
' that setup can run to completion.

Print #ff, "           </FirstLogonCommands>"
Print #ff, "           <AutoLogon>"
Print #ff, "               <Password>"
Print #ff, "                   <Value>Password1</Value>"
Print #ff, "                   <PlainText>true</PlainText>"
Print #ff, "               </Password>"
Print #ff, "               <Username>"; UserName$; "</Username>"
Print #ff, "               <Enabled>true</Enabled>"
Print #ff, "               <LogonCount>1</LogonCount>"
Print #ff, "           </AutoLogon>"
Print #ff, "       </component>"
Print #ff, "   </settings>"

' Start the Specialize pass

Print #ff, "   <settings pass="; Chr$(34); "specialize"; Chr$(34); ">"
Print #ff, "       <component name="; Chr$(34); "Microsoft-Windows-Shell-Setup"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); ">"

' Set the computer name. If no name was provided, then setup will assign a random name.

Print #ff, "           <ComputerName>"; ComputerName$; "</ComputerName>"

' Set time zone

Print #ff, "           <TimeZone>"; TimeZone$; "</TimeZone>"

' Create the Windows Deployment section

Print #ff, "       </component>"
Print #ff, "       <component name="; Chr$(34); "Microsoft-Windows-Deployment"; Chr$(34); " processorArchitecture="; Chr$(34); "amd64"; Chr$(34); " publicKeyToken="; Chr$(34); "31bf3856ad364e35"; Chr$(34); " language="; Chr$(34); "neutral"; Chr$(34); " versionScope="; Chr$(34); "nonSxS"; Chr$(34); " xmlns:wcm="; Chr$(34); "http://schemas.microsoft.com/WMIConfig/2002/State"; Chr$(34); " xmlns:xsi="; Chr$(34); "http://www.w3.org/2001/XMLSchema-instance"; Chr$(34); ">"

' Use a command to bypass quality updates during setup. Use another command to prevent auto device encryption.
' First, we create the RunSynchronous block so that we can add those commands.

Print #ff, "           <RunSynchronous>"

' Create a counter called Phase4CommandCounter to keep track of command numbers. Increment every
' time a command is added.

Phase4Commandcounter = 1

' This command disables networking so that quality updates cannot be installed during setup. As a result, add it only
' if quality updates are being bypassed.

If SystemType$ = "BIOS" GoTo DeviceEncryption
If BypassQualityUpdatesDuringOobe$ = "N" GoTo DeviceEncryption

Print #ff, "               <RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <Order>"; LTrim$(Str$(Phase4Commandcounter)); "</Order>"
Print #ff, "                   <Path>powershell.exe -Command &quot;Get-NetAdapter | ForEach-Object { Disable-NetAdapter -Name $_.Name -Confirm:$false }&quot;</Path>"
Print #ff, "               </RunSynchronousCommand>"

Phase4Commandcounter = Phase4Commandcounter + 1

DeviceEncryption:

' This command bypasses automatic device encryption. Add it only if user elected to bypass device encryption.

If SystemType$ = "BIOS" GoTo DevEncryptDone
If BypassDeviceEncryption$ = "N" GoTo DevEncryptDone

Print #ff, "               <RunSynchronousCommand wcm:action="; Chr$(34); "add"; Chr$(34); ">"
Print #ff, "                   <Order>"; LTrim$(Str$(Phase4Commandcounter)); "</Order>"
Print #ff, "                   <Path>reg add HKLM\System\CurrentControlSet\Control\BitLocker /v PreventDeviceEncryption /t reg_dword /d 0x00000001 /f</Path>"
Print #ff, "               </RunSynchronousCommand>"

Phase4Commandcounter = Phase4Commandcounter + 1

DevEncryptDone:

' Close the RunSynchronous block

Print #ff, "           </RunSynchronous>"

' Close everything else

Print #ff, "       </component>"
Print #ff, "   </settings>"
Print #ff, "</unattend>"

' Close the autounattend.xml file

Close #ff

Cls
Print "All operations have been completed and the autounattend.xml can be found in the same location as this program."
Print "Here are some notes about the options that you selected:"
Print
If UserLocale$ = "en-001" Then
    Color 14, 4: Print "IMPORTANT:";: Color 15: Print " You have chosen to use ";: Color 10: Print "en-001";: Color 15: Print " for the UserLocale. By using this option, you will need to take some additional"
    Print "           steps after Windows is installed. It is strongly suggested that you do this right away to avoid forgetting."
    Print
    Print "Step 1: Open Start. Do you see any greyed out placeholders in the empty positions where icons would be located? If so,"
    Print "        this means that the system has not had access to the Internet yet and this is absolutely necessary before you"
    Print "        move to step 2. Enable networking now and configure your machine so that it can reach the Internet. Once the"
    Print "        placeholders on the Start screen have disappeared, you can then move on tho step 2. This should happen almost"
    Print "        immediately upon getting Internet access."
    Print
    Print "Step 2: Go to Start > Time & language > Language & region. Change ";: Color 10: Print "Country or region";: Color 15: Print " to ";: Color 10: Print "United States";: Color 15: Print "."
    Print "        In addition, change ";: Color 10: Print "Regional format";: Color 15: Print " to ";: Color 10: Print "English (United States)";: Color 15: Print "."
    Print
    Print "This completes the procedure."
    Pause
End If

Cls
Color 0, 10: Print "Thanks for using the Answer File Generator!": Color 15
Print
Print "Pressing any key will exit the program."

GoTo EndProgram


' ****************
' * Program Help *
' ****************

' This section provides comprehensive help for this program. It includes general program help
' and detailed information for each feature of the program.

ProgramHelp:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " Dual Architecture Edition         "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 50
Print " Program Help Menu ";
Locate 7, 1
Color 0, 14
Print "    1) Inject Windows updates into one or more Windows editions and create a multi edition bootable image       "
Print "    2) Inject drivers into one or more Windows editions and create a multi edition bootable image               "
Print "    3) Inject boot-critical drivers into one or more Windows editions and create a multi edition bootable image "
Print "    4) Modify Windows ISO to bypass system requirements and optionally force use of previous version of setup   "
Color 0, 10
Print "    5) Make or update a bootable drive from one or more Windows / WinPE / WinRE ISO images                      "
Print "    6) Create a bootable Windows ISO image that can include multiple editions                                   "
Print "    7) Create a bootable ISO image from Windows files in a folder                                               "
Print "    8) Reorganize the contents of a Windows ISO image                                                           "
Print "    9) Convert between an ESD and WIM either standalone or in an ISO image                                      "
Color 0, 3
Print "   10) Get image info - display basic info for each edition in an ISO image and display Windows build number    "
Print "   11) Modify the NAME and DESCRIPTION values for entries in a WIM file                                         "
Color 0, 6
Print "   12) Export drivers from this system                                                                          "
Print "   13) Expand drivers supplied in a .CAB file                                                                   "
Print "   14) Create a Virtual Disk (VHDX) - NOTE: Win 11 23H2+ has a new GUI to make doing this from the OS easy      "
Print "   15) Create a VHD, deploy Windows to it, and add it to the boot menu to make a dual boot configuration        "
Print "   16) Create a generic ISO image and inject files and folders into it                                          "
Print "   17) Cleanup files and folders                                                                                "
Color 0, 55
Print "   18) Unattended answer file generator                                                                         "; ""
Color 0, 8
Print "   19) Exit                                                                                                     "
Color 0, 13
Print "   20) Get general help on the use of this program                                                              "
Color 0, 15
Print "   21) Exit help and return to main menu                                                                        "
Locate 29, 0
Color 15
Input "   Please select the item you would like help with by entering its number (21 returns to the main menu): ", MenuSelection

Select Case MenuSelection
    Case 1
        GoTo HelpInjectUpdates
    Case 2
        GoTo HelpInjectDrivers
    Case 3
        GoTo HelpInjectBCD
    Case 4
        GoTo HelpBypassWin11Requirements
    Case 5
        GoTo HelpMakeMultiBootImage
    Case 6
        GoTo HelpMakeBootDisk2
    Case 7
        GoTo HelpCreateBootableISOFromFiles
    Case 8
        GoTo HelpChangeOrder
    Case 9
        GoTo HelpConvertEsd
    Case 10
        GoTo HelpGetWimInfo
    Case 11
        GoTo HelpNameAndDescription
    Case 12
        GoTo HelpExportDrivers
    Case 13
        GoTo HelpExpandDrivers
    Case 14
        GoTo HelpCreateVHDX
    Case 15
        GoTo HelpAddVHDtoBootMenu
    Case 16
        GoTo HelpCreateISOImage
    Case 17
        GoTo HelpGetFolderToClean
    Case 18
        GoTo HelpAnswerFileGen
    Case 19
        GoTo HelpExit
    Case 20
        GoTo GeneralHelp
    Case 21
        ChDir ProgramStartDir$: GoTo BeginProgram
End Select

' We arrive here if the user makes an invalid selection from the main menu

Cls
Color 14, 4
Print
Print "You have made an invalid selection.";
Color 15
Print " You need to make a selection by entering a number from 1 to 21."
Pause
GoTo ProgramHelp

' Help Topic: Get general help on the use of this program

GeneralHelp:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - General Help ";
Locate 9, 1
Color 15
Print "    1) System requirements       "
Print "    2) Terminology"
Print "    3) Responding to the program"
Print "    4) Hard disk vs removable media"
Print "    5) Reviewing log files"
Print "    6) How answer files are handled"
Print "    7) Auto shutdown, hibernation, and program pause"
Print "    8) Antivirus exclusions"
Print
Color 0, 13
Print "    9) Return to main help menu "
Locate 28, 0
Color 15
Input "   Please select the item you would like help with by entering its number (9 returns to the main help menu): ", MenuSelection

Select Case MenuSelection
    Case 1
        GoTo SystemRequirements
    Case 2
        GoTo Terminology
    Case 3
        GoTo Responding
    Case 4
        GoTo MediaTypes
    Case 5
        GoTo LogFileReview
    Case 6
        GoTo AnswerFiles
    Case 7
        GoTo ShutdownAndPause
    Case 8
        GoTo AV_Exclusions
    Case 9
        GoTo ProgramHelp
End Select

' We arrive here if the user makes an invalid selection from the main menu

Color 15
Cls
Color 14, 4
Print
Print "You have made an invalid selection.";
Color 15
Print " You need to make a selection by entering a number from 1 to 9."
Pause
GoTo GeneralHelp

' Help Topic: Get general help on the use of this program > System requirements

SystemRequirements:

Color 15
Cls
Print "System Requirements"
Print "==================="
Print
Print "This program is a 64-bit program designed to work only on 64-bit systems. It is also designed to work with Windows ISO"
Print "images that contain an install.wim file in the \sources folder, not an INSTALL.ESD. There are a couple of exceptions"
Print "which are addressed in the help sections related to those sections where applicable."
Print
Print "This program requires the Windows ADK to be installed. Only the ";: Color 0, 10: Print "Deployment Tools";: Color 15: Print " component needs to be installed. The"
Print "program will display a warning when it is started if the ADK is not installed. However, it will continue to operate"
Print "since some functions will work without the ADK. If the user selects a feature from the menu that requires the ADK, the"
Print "user will be warned and returned to the main menu."
Print
Print "Run the program locally, not from a network location."
Print
Print "When operating on multiple editions of Windows in the same project (for example, Win 10 Pro, Home, Education editions,"
Print "etc.), this program is designed to work with editions of the same version. For example, you do not want to mix version"
Print "21H1 and 21H2 in the same project. Do ";: Color 0, 10: Print "NOT";: Color 15: Print " create ISO images that have Windows editions having different builds."
Print
Print "Disable QuickEdit Mode - The color of text in some places within the program may be displayed incorrectly if QuickEdit"
Print "mode is enabled. You should disable QuickEdit mode by following these steps: Right-click on the title bar when the"
Print "program is running and choose Properties, select the Options tab, uncheck QuickEdit Mode, click OK, restart the program."
Print "You should only need to do this once as this setting will be retained for the next time you run the program."
Pause
GoTo GeneralHelp

' Help Topic: Get general help on the use of this program > Terminology

Terminology:

Color 15
Cls
Print "Terminology - Images, Editions, and Indices"
Print "==========================================="
Print
Print "The Windows "; Chr$(34); "image"; Chr$(34); " files distributed by Microsoft are known as "; Chr$(34); "ISO image"; Chr$(34); " files. These images hold a collection of"
Print "files in some ways like a ZIP file. These images, in turn, hold other images, known as WIM or Windows IMage files. These"
Print "image files hold information regarding one or more flavors of Windows such as Pro, Home, Education, etc. These various"
Print "flavors of windows are also known as "; Chr$(34); "editions"; Chr$(34); ". Each edition of Windows has an"; Chr$(34); "index"; Chr$(34); " number associated with it."
Print "Windows utilities work with the index number that is associated with each edition of Windows. When an index number is"
Print "needed, this program will assist you in determining the index number associated with the editions of Windows that you"
Print "want to work with."
Print
Print "Please also note that throughout the program we use the terms "; Chr$(34); "Partition"; Chr$(34); " and "; Chr$(34); "Volume"; Chr$(34); " interchangably. While there are"
Print "technically differences, please note that we are not differentiating between these here."
Pause
GoTo GeneralHelp

' Help Topic: Get general help on the use of this program > Responding to the program

Responding:
Cls
Print "Responding to the Program"
Print "========================="
Print
Print "When the program asks for a path, you can enclose paths that have spaces in quotes if you wish, but this is not"
Print "necessary. The program will handle paths either way. Make sure that paths and filenames do NOT contain commas."
Print
Print "The program will allow you to paste into it, which can be especially helpful for long paths. To copy and paste a path"
Print "do the following:"
Print
Print "  Locate the path in File Explorer."
Print "  With the folder of interest open in File Explorer, right-click in the Address Bar."
Print "  Select "; Chr$(34); "Copy address"; Chr$(34); "."
Print "  In the WIM Tools program, right-click to paste the address."
Print
Print "If you want to copy and paste a path including the file name, follow these steps:"
Print
Print "  In File Explorer, open the folder that contains the file of interest."
Print "  In the right pane, while holding down the SHIFT key, right-click on the file name."
Print "  From the menu that opens, choose "; Chr$(34); "Copy as path"; Chr$(34); "."
Print "  In the WIM Tools program, right-click to paste the path with the file name."
Print
Print "NOTE: In Windows 11, there is no need to hold down the SHIFT key. Simply right-click and choose "; Chr$(34); "Copy as path"; Chr$(34); "."
Print
Print "Yes or No Responses - Respond with anything that starts with a "; Chr$(34); "Y"; Chr$(34); " or "; Chr$(34); "N"; Chr$(34); " in uppercase, lowercase, or mixed case."
Pause
Cls
Print "Scripting Responses"
Print "==================="
Print
Print "The program maintains a keyboard buffer. As a result, a series of responses can simply be pasted into the program."
Print
Color 0, 10: Print "IMPORTANT:";: Color 15: Print " Because the program maintains a keyboard buffer, you should ";: Color 0, 10: Print "NOT";: Color 15: Print " press keys at random while the program is"
Print "performing an operation and has focus. As soon as an operation is completed the keys you pressed will be processed."
Print
Print "However, the program now has a powerful script recording and playback tool. When you select the option to inject"
Print "Windows updates, drivers, or boot-critical drivers, you will be given the option to record or playback a script. The"
Print "scripts that you record will be saved to the folder in which the program is located as ";: Color 0, 10: Print "WIM_SCRIPT.TXT";: Color 15: Print ". If you wish to"
Print "record multiple scripts, make sure to rename your scripts to avoid overwriting them. Note that the scripts will include"
Print "comments to make it easy for you to manually alter the scripts. Comments begin with two colons (::). You can add"
Print "comments of your own if you wish. Any line that does not start with two colons is an actual response typed by the user."
Print
Print "Note that you do not have to execute an injection of Windows updates, drivers, or boot-critical drivers to create the"
Print "script. Once all the needed information to create the script has been gathered, the program will ask if you wish to"
Print "proceed and perform the updates. If not, the script will still be saved without performing any of the updates."
Print
Color 0, 10: Print "Use caution with scripting!";: Color 15: Print " If you change the contents of folders you are working with, the script may not work."
Print "A new feature of the program is the ability to specify individual file names to be updated rather than a folder and"
Print "having the program ask if you want to update each file. By specifically using the full path with file name in your"
Print "responses, the script will be more reliable. This is because when a folder is chosen, your responses change"
Print "depending upon how many files are in the folder. If that changes, your responses won't work when played back. When not"
Print "scripting, it's easier to point to a folder, but for scripting, it's more reliable to take the time to specify each"
Print "file to update. This will also make manually modifying the script responses easier."
Print
Pause
Cls
Print "Entering Index Numbers"
Print "======================"
Print
Print "Some Windows ISO images may have multiple editions of Windows. Each edition has a unique index number."
Print
Print "Some of the routines in WIM Tools will ask for the index number(s) to work with. These routines will have an option to"
Print "allow you to display all the available editions of Windows along with the indices associated with each edition."
Print "Index numbers can be entered into the program as shown in the examples below:"
Print
Print "- A single index number. Example: 6"
Print
Print "- Multiple index numbers: Separate indices with a space. Example: 1 6 9"
Print
Print "- A range of indices: Specify a range of indices by separating the lower and upper range with a dash. Example: 1-3"
Print
Print "- A combination of the above: Separate each group with a space. Example: 1 4-6 9"
Print
Print "- The word "; Chr$(34); "ALL"; Chr$(34); ". This will automatically determine all available indices and select them."
Print
Color 0, 10: Print "IMPORTANT:";: Color 15: Print " Enter the index numbers from low to high. Don't specify a lower index number after a higher index number."
Print "The exception to this is the routine for reorganizing the contents of a Windows ISO image. Since the goal of this"
Print "option is to reorder the Windows editions, it is okay to specify indices in any order."
Pause
GoTo GeneralHelp

' Help Topic: Get general help on the use of this program > Hard disk vs removable media

MediaTypes:

Cls
Print "Hard Disk vs Removable Media"
Print "============================"
Print
Print "For the process of injecting Windows updates or drivers, one of the Microsoft utilities used (DISM) requires the use of"
Print "a fixed disk, not a removable disk. Make sure that you use a fixed disk to store your projects. You can work around"
Print "this by creating a virtual hard disk on a removable disk and then using that virtual hard disk to store your project."
Print "There is an option on the main menu to help you create a virtual disk should you need to do so."
Pause
GoTo GeneralHelp

' Help Topic: Get general help on the use of this program > Reviewing log files

LogFileReview:

Cls
Print "Reviewing Log Files"
Print "==================="
Print
Print "After running routines that inject files into the Windows image, you should review log files to check for errors."
Print
Print "You will find a log named "; Chr$(34); "PendingOps.log"; Chr$(34); ". This log file will show any pending operations that were set to take place."
Print "Pending opertions occur when certain updates are applied. As an example, if you enable NetFX3, your Windows edition will"
Print "have a Pending Install operation. This will prevent DISM Image Cleanup operations from being able to run. The program"
Print "will display a warning about this in the header if this condition is detected. Note that the program will continue to"
Print "run, but you are probably best served by reapplying updates to images where Windows editions do not have pending"
Print "operations. After the routine completes, you will also find a log file in the logs folder named "; Chr$(34); "PendingOps.log"; Chr$(34); "."
Print "This log file will tell you what source file and index number(s) in that source had the pending operations."
Print
Print "In addition, you will find log files with names such as x64_1_UpdateResults.txt. These log files will allow you to see"
Print "what updates have been applied to your WIM images. One such log will be present for each edition present. Note that the"
Print "number after the underscore will match the index number of that Windows edition in the final image."
Pause
GoTo GeneralHelp

' Help Topic: Get general help on the use of this program > How answer files are handled

AnswerFiles:

Cls
Print "How Answer Files are Handled"
Print "============================"
Print
Print "Including an answer file in your project can cause the system or VM booted from an image or bootable media to begin a"
Print "fully automated installation of Windows that can result in one or more disks being wiped without warning. As a result,"
Print "all routines in this program that can compile multiple Windows editions into a single image or bootable media will"
Print "exclude the answer file from the original sources. One reason for this is that it is possible to have different answer"
Print "files on different sources and as a result the final copied answer file could potentially be from any of these sources,"
Print "possibly an answer file you did not mean to use. There are 2 exceptions to this:"
Print
Print "  1) For the routine to Make or update a bootable drive from one or more Windows / WinPE / WinRE ISO images, we will"
Print "     ask the user if they want to exclude an answer file if one exists."
Print
Print "  2) For the routine that creates a bootable ISO image from files in a folder we will INCLUDE the answer file if it is"
Print "     present since the user can simply delete it from the folder before creating the image if it is not wanted."
Print
Print "Note that for the routine that injects Windows updates, it will exclude answer files from the original sources, but if"
Print "you include an answer file in the "; Chr$(34); "Answer_File"; Chr$(34); " subfolder of the folder that you specify as the Windows updates"
Print "location, this answer file will be copied since you are specifically including it in the update files."
Print
Print "Please see the help topic "; Chr$(34); "Inject Windows updates into one or more Windows editions and create a multi edition"
Print "bootable image"; Chr$(34); " > "; Chr$(34); "Organizing your updates"; Chr$(34); " for information on adding an answer file."
Print
Print "Note that you can always extract the contents of an ISO image to a folder on your hard disk, add your answer file to"
Print "this folder and finally use the routine "; Chr$(34); "Create a bootable ISO image from Windows files in a folder"; Chr$(34); " to create a new"
Print "image that includes the answer file."
Pause
GoTo GeneralHelp

' Help Topic: Get general help on the use of this program > Auto shutdown and program pause

ShutdownAndPause:

Cls
Print "Auto Shutdown and Program Pause"
Print "==============================="
Print
Print "For the routines that inject Windows updates, drivers, or boot critical drivers, you can choose to have the program"
Print "automatically shutdown or hibernate the system when it is done running. You can also pause the execution of the program"
Print "to free up resources to other programs."
Print
Print "To have the program perform an automatic SHUTDOWN, create a file on your desktop named "; Chr$(34); "Auto_Shutdown.txt"; Chr$(34); ". If a file by"
Print "that name exists, a shutdown will be performed when the program is finished. Use "; Chr$(34); "Auto_Hibernate.txt"; Chr$(34); " if you would"
Print "prefer hibernation. Having both files at the same time is invalid and then neither action will take place."
Print
Print "To pause program execution, create a file on your desktop named "; Chr$(34); "WIM_Pause.txt"; Chr$(34); ". If a file by that name exists, the"
Print "program will pause at the next step in the update process. Note that you can have both a Auto_Shutdown.txt file or a"
Print "Auto_Hibernate.txt file along with a WIM_Pause.txt file."
Print
Print "Note that for auto shutdown and hibernation, you can change your mind at any time. The existence of this file will only"
Print "take effect when the routine finishes running so you can create the file even while the program is running, or you can"
Print "delete / rename the file if you decide at some point that you do not want an automatic shutdown / hibernation."
Print
Print "While the program is running, you will see a status indication on the upper right of the screen to remind you of the"
Print "current status. If program execution is paused, a flashing message will be displayed as a reminder that the program is"
Print "paused. Note that when you make a change, the status will not update immediately. The status is updated each time the"
Print "program advances to the next item in the displayed checklist. However, when the program is resumed by deleting or"
Print "renaming the "; Chr$(34); "WIM_Pause.txt"; Chr$(34); " file, this status change will be reflected immediately."
Pause
Cls
Print "Note that when the system is shutdown or hibernated automatically, any status messages or warning that would normally be"
Print "displayed will not be shown due to the shutdown or hibernation. Instead, this information is logged to a file. The next"
Print "time the program is run, that information will be automatically displayed. After viewing this information, that file"
Print "will be automatically deleted."
Print
Print "Note that none of these filenames are case sensitive."
Pause
GoTo GeneralHelp

' Help Topic: Get general help on the use of this program > Antivirus exclusions

AV_Exclusions:

Cls
Print "Antivirus Exclusions"
Print "===================="
Print
Print "When the program is started, it will set an exclusion for itself in Windows Defender Antivirus. In addition, for the"
Print "projects that inject files, drivers, or boot-critical drivers, the program will set a Windows Defender Antivirus"
Print "exclusion for the destination folder to improve performance."
Print
Print "When choosing "; Chr$(34); "Exit"; Chr$(34); " from the Main Menu, the exclusions will be removed. Note that if the program ends without"
Print "specifically choosing "; Chr$(34); "Exit"; Chr$(34); " from the main menu, the exclusions will remain in place. To remove the exclusion, simply"
Print "start the program, then choose "; Chr$(34); "Exit"; Chr$(34); " from the Main Menu. The exclusions will be removed."
Print
Print "To manually check or manage exclusions, press Windows_Key + R and type in "; Chr$(34); "ms-settings:windowsdefender"; Chr$(34); ". Press ENTER. Go"
Print "to Virus & threat protection. Under the section Virus & threat protection settings choose Manage settings >  Add or"
Print "remove exclusions."
Print
Print "If you use other AV software, you can set exclusions manually if you wish."
Pause
GoTo GeneralHelp

' Help Topic: Inject Windows updates into one or more Windows editions and create a multi edition bootable image

HelpInjectUpdates:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Inject Windows updates into one or more Windows editions";
Locate 4, 38
Print "                and create a multi edition bootable image               ";
Locate 9, 1
Color 15
Print "    1) General information about this routine"
Print "    2) Acceptable Windows images"
Print "    3) How to obtain updates and select the correct update packages"
Print "    4) Working with dual architecture images (applies only to dual architecture edition of this program)"
Print "    5) Organizing update files"
Print
Color 0, 13
Print "    6) Return to main help menu "
Locate 28, 0
Color 15
Input "   Please select the item you would like help with by entering its number (6 returns to the main help menu): ", MenuSelection

Select Case MenuSelection
    Case 1
        GoTo InjectUpdatesGeneralInfo
    Case 2
        GoTo AcceptableImages
    Case 3
        GoTo ObtainingUpdates
    Case 4
        GoTo WorkWithDualArcImages
    Case 5
        GoTo OrganizingUpdates
    Case 6
        GoTo ProgramHelp
End Select

' We arrive here if the user makes an invalid selection from the main menu

Color 15
Cls
Color 14, 4
Print
Print "You have made an invalid selection.";
Color 15
Print " You need to make a selection by entering a number from 1 to 6."
Pause
GoTo HelpInjectUpdates

' Help Topic: Inject Windows updates into one or more Windows editions and create a multi edition bootable image > General information about this routine

InjectUpdatesGeneralInfo:

Cls
Print "General Information About This Routine"
Print "======================================"
Print
Print "This routine will allow the user to pick Windows editions, inject Windows updates into them, and then create a single"
Print "ISO image containing all these Windows editions. One or more Windows editions from an ISO image can be selected, as"
Print "well as Windows editions from multiple different ISO images. A mix of both x86 and x64 images can be used in the same"
Print "project (applies only to the Dual Architecture edition of this program). Even sysprep images can be added."
Print
Print "This process will NOT inject driver updates. This routine only injects Windows updates. There is a separate option from"
Print "the main menu that will inject drivers."
Pause
GoTo HelpInjectUpdates

' Help Topic: Inject Windows updates into one or more Windows editions and create a multi edition bootable image > Acceptable Windows images

AcceptableImages:

Cls
Print "Acceptable Windows Images"
Print "========================="
Print
Print "All Windows ISO images used need to have an install.wim (not an INSTALL.ESD). The one exception to this is that the file"
Print "used to build the base image for a dual architecture project (one that contains both x64 and x86 images), can have"
Print "INSTALL.ESD files on it. Please see the help topic Working with Dual Architecture Images for more information."
Pause
GoTo HelpInjectUpdates

' Help Topic: Inject Windows updates into one or more Windows editions and create a multi edition bootable image > How to obtain updates and select the correct update packages

ObtainingUpdates:
Cls
Print "How to Obtain Updates and Select the Correct Update Packages"
Print "============================================================"
Print
Print "Go to the Microsoft Update Catalog (https://www.catalog.update.microsoft.com/) to obtain updates. Search for "; Chr$(34); "Windows"
Print "11 Version 23H2"; Chr$(34); " (or whatever your version is) and then click on the Last Updated column to sort with the latest"
Print "updates first."
Print
Print "For the MicroCode updates you will need to search for these seperately through a search engine such as Google or Bing."
Print
Print "Make sure to download the correct updates (either x64 or x86). x86 updates will never be needed in the x64 Only Edition"
Print "of this program."
Print
Print "IMPORTANT: When downloading updates from the Microsoft Update Catalog, you will see that some updates are described as"
Print "dynamic updates. The Safe OS Dynamic Update and the Setup Dynamic Update are available only as dynamic updates. However,"
Print "what may be confusing is that the Cumulative Update may be available as both a standard update AND a dynamic update."
Print "For the purposes of this program, ";: Color 0, 10: Print "DO NOT";: Color 15: Print " use the dynamic version of the cumulative update. Instead, download the version"
Print "that is not described as a dynamic update."
Pause
GoTo HelpInjectUpdates

' Help Topic: Inject Windows updates into one or more Windows editions and create a multi edition bootable image > Working with dual architecture images (applies only to dual architecture edition of this program)

WorkWithDualArcImages:

Cls
Print "Working with Dual Architecture Images (Applies Only to Dual Architecture Edition of This Program)"
Print "================================================================================================="
Print
Print "When creating a base image for a dual architecture project with both x86 and x64 editions of Windows, neither an x86 nor"
Print "an x64 image contain all the files that we need to create the base image. If any additional files are needed the program"
Print "will prompt for the location of either a dual architecture image or a "; Chr$(34); "Windows Boot Foundation"; Chr$(34); " image."
Print
Print "Option 1: Download a dual architecture image from the Windows Media Creation Tool Web Site here:"
Print
Print "https://www.microsoft.com/en-us/software-download/windows10"
Print
Print "   Select Download tool now and run it."
Print "   When the tool runs, accept the license terms."
Print "   Choose Create installation media (USB flash drive, DVD, or ISO file) for another PC and then click on Next."
Print "   Uncheck the checkbox for Use the recommended options for this PC."
Print "   Select the appropriate language."
Print "   For Architecture make certain to select Both."
Print "   Click on Next."
Print "   Select ISO file and then Next."
Print "   Save the downloaded image."
Pause
Cls
Print "Option 2: To create a Windows Boot Foundation ISO image perform these steps:"
Print
Print "The advantage of this method is that a Boot Foundation ISO image is small, less than 50 MB. This means that the huge"
Print "dual architecture ISO image does not need to be kept."
Print
Print "   Download the dual architecture image as noted in option 1"
Print "   Copy all files and folders EXCEPT the x64 and x86 folder to a temporary location."
Print "   Create an ISO image from the folders and files in the temporary location (this program has a routine to create a"
Print "   generic ISO image that can be used for this). Make sure to put these files and folders at the root of the ISO image."
Print
Print "That ISO image is your Boot Foundation ISO image. Note that "; Chr$(34); "Boot Foundation Image"; Chr$(34); " is a term I have created."
Print
Print "At this point the dual architecture image is no longer needed. Feel free to discard it."
Print
Print "TIP: The routine to inject updates into one or more Windows editions will check each edition as it is processed for any"
Print Chr$(34); "Pending Operations"; Chr$(34); "."; " As an example, if you enable NetFX3, your Windows edition will have a Pending Install"
Print "operation. This will prevent DISM Image Cleanup operations from being able to run. The program will display a warning"
Print "about this in the header if this condition is detected. Note that the program will continue to run, but you are probably"
Print "best served by reapplying updates to images where Windows editions do not have pending operations. After the routine"
Print "completes, you will also find a log file in the logs folder named PendingOps.log. This log file will tell you what"
Print "source file and index number(s) in that source had the pending operations."
Pause
GoTo HelpInjectUpdates

' Help Topic: Inject Windows updates into one or more Windows editions and create a multi edition bootable image > Organizing your update files
OrganizingUpdates:

Cls
Print "Organizing Update Files"
Print "======================="
Print
Print "The program will expect to find Windows updates in a particular folder structure. Create a folder in which to store all"
Print "update files (D:\Win 10 Update Files in this example), then create subfolders that look like this:"
Print
Print "D:\Win 10 Update Files"
Print "      \Answer_File      < Place autounattend.xml answer file here if you want to include one"
Print "      \x64              < Folders with x64 (64-bit) drivers will be organized under this folder"
Print "          \SSU          < For those rare instances when a Standalone SSU is available"
Print "          \LCU          < Place the LCU (Latest Cumulative Update) in this folder (see note 2 pages down)"
Print "          \Other        < Other updates such as .NET update or OOBE ZDP updates"
Print "          \PE_Files     < Used to copy generic files (not Windows updates) to the Win PE image"
Print "          \SafeOS_DU    < Safe OS Dynamic Update - Used to update Windows Recovery (WinRE)"
Print "          \Setup_DU     < Setup Dynamic Update"
Print "      \x86              < Folders with x86 (32-bit) drivers will be found here"
Print "          \SSU"
Print "          \LCU"
Print "          \Other"
Print "          \PE_Files"
Print "          \SafeOS_DU"
Print "          \Setup_DU"
Print
Print "The next page will display some more information about the above folder structure. We will include the above list of"
Print "folders so that you can refer to it easily along with the explanation."
Pause
Cls
Print "D:\Win 10 Update Files"
Print "      \Answer_File      < Place autounattend.xml answer file here if you want to include one"
Print "      \x64              < Folders with x64 (64-bit) drivers will be organized under this folder"
Print "          \SSU          < For those rare instances when a Standalone SSU is available"
Print "          \LCU          < Place the LCU (Latest Cumulative Update) in this folder (see note on next page)"
Print "          \Other        < Other updates such as .NET update or OOBE ZDP updates"
Print "          \PE_Files     < Used to copy generic files (not Windows updates) to the Win PE image. RARELY needed."
Print "          \SafeOS_DU    < Safe OS Dynamic Update - Used to update Windows Recovery (WinRE)"
Print "          \Setup_DU     < Setup Dynamic Update"
Print "      \x86              < Folders with x86 (32-bit) drivers will be found here"
Print "          \SSU"
Print "          \LCU"
Print "          \Other"
Print "          \PE_Files"
Print "          \SafeOS_DU"
Print "          \Setup_DU"
Print
Print "In the folder tree above, we have 3 folders in the D:\Win 10 Updates folder. We have Answer_File, x64, and x86."
Print "Because both x86 and x64 versions of Windows can have all the same kinds of updates, we need to keep them separate so"
Print "that the program can pick the correct updates for the Windows image(s) with which it is working. Note that if you are"
Print "using only one kind of image, for example, only x64 images, there is no need to create the other folder such as x86"
Print "or anything beneath it, but having it present will not affect operation. Since Win 11 is only x64, you will never need"
Print "the x86 folder in Win 11 projects. However, it won't cause any problems if it is present."
Pause
Cls
Print "A note about the LCU folder:"
Print
Print "Prior to Windows 11 24H2, the main OS was updated with cumulative updates. Windows 11 24H2 introduces checkpoint and"
Print "incremental updates. You can recognize these because the Microsft update catalog will show more than one downloadable"
Print "file for a cumulative update if an incremental update is available. If only one file is shown, this is a checkpoint"
Print "update. If two files are shown, this includes a checkpoint update and an incremental update. Download both and save"
Print "them both in the LCU folder. If only one downloadable file is shown, just download that one file and save it in the"
Print "LCU folder. This one file will be a cumulative update (prior to Win 11 24H2) or a checkpoint update (Win 11 24H2+."
Pause
Cls
Print "Description of Each Folder in the Structure"
Print "==========================================="
Print
Print "Answer_File - An autounattend.xml file will NOT be included from your original Windows image file(s). If you want to"
Print "include one in the final image, place it into the Answer_File folder. We do this for a reason: Let's say you have"
Print "multiple Windows images that have an answer file. There is no way for the program to know which one to use. Also, it"
Print "is a potential safety issue if you end up with an unexpected answer file in your image and you then boot from that"
Print "image. This could cause an uxpected loss of data as Windows starts installing unexpectly and wipes your drive. By"
Print "doing this, you are explicitly adding an unattended installation answer file to your image."
Print
Print "SSU - Normally, the SSU is included in the LCU package, however, there may be rare circumstances where Microsoft"
Print "will release a Standalone SSU. If available, place a Standalone SSU into this folder."
Print
Print "LCU - Installs the ";: Color 0, 10: Print "L";: Color 15: Print "atest ";: Color 0, 10: Print "C";: Color 15: Print "umulative quality ";: Color 0, 10: Print "U";: Color 15: Print "pdate. When downloading this update from the Microsoft Update Catalog,"
Print "the Title field in the update catalog will show Cumulative Update for Windows 10 (or 11) for this update. Store only"
Print "one LCU file in this folder except for Checkpoint / Incremental updates. For these, save the most recent checkpoint"
Print "and incremental update as noted on the previous pages of this help topic."
Print
Print "Other - Updates that do not fall into the category of any other folder here are placed into the \Other folder. These"
Print "typically include .NET updates as an example. You can place multiple files in this folder, but you should save only the"
Print "latest version of each update type here. There is one exception to this rule: For OOBE ZDP updates, you should"
Print "download ALL available updates of this type becasue these are not cumulative updates. Note that Microcode updates can"
Print "also be placed here."
Pause
Cls
Print "PE_Files - THIS IS AN ADVANCED FEATURE that few people will ever need. No Microsoft update files are placed into this"
Print "folder. You can leave this folder empty or omit the folder entirely. Use this folder to keep files that need to be"
Print "accessible to Windows PE during Windows setup in this folder. This would typically include things like scripts. Simply"
Print "place any such files in this folder. If you later run Windows setup from media created with these files included, you"
Print "will find the files that you placed here on X:\. This is the RAM Drive that Windows setup creates during installation."
Print "To delete a file from the WinPE image, create a dummy file with the same name as the file that you want to delete"
Print "preceded with a minus sign (-). For example, to delete a file called MyScript.bat, create any file in the \PE_Files"
Print "folder, then rename the file to -MyScript.bat. The case of the filename does not matter. You should rarely, if ever,"
Print "need to use this."
Print
Print "SafeOS_DU - Fixes for the Safe OS that are used to update Windows recovery environment (WinRE). Save your Safe OS"
Print "Dynamic Update file here. When downloading files from the Microsoft Update Catalog, the Safe OS Dynamic Update will"
Print "specifically indicate Safe OS Dynamic Update in the Title field."
Print
Print "Setup_DU - Fixes to Setup binaries or any files that Setup uses for feature updates. Note that when downloading files"
Print "from the Microsoft Update Catalog, the Setup Dynamic Update will indicate Windows 10 (or 11) Dynamic Update in the"
Print "Product field and Dynamic Update for Windows 10 (or 11) in the Title field. Store only one Setup Dynamic Update file"
Print "in this folder."
Print
Print "TIP: If you have a new update that you wish the apply to your Windows edition(s) and you already have other updates"
Print "applied to your Windows edition(s), create an update folder with only the new update(s) and remove the contents of all"
Print "the other folders. This will cause the updates to be applied faster since the other updates don't have to be parsed to"
Print "see if their contents have already been applied. Note that this is not mandatory. As an example, suppose that you have"
Print "already applied all of the updates for the month to a Windows image. Then, Microsoft releases an out-of-band emergency"
Print "safe OS Dynamic Update. You can copy your updates folder and delete the LCU folder (or just the contents), as well as"
Print "all of the other folders EXCEPT the SafeOS_DU folder under x64 or x86 to add just this one update. This not mandatory."
Print "It is fine to leave updates you have already installed in place, it may simply take longer to process."
Pause
Cls
Print "Final Tip: While some updates, such as the Latest Cumulative Update (LCU), are released every month, some updates are"
Print "not released every month. You may need to go back in time to find the latest release of a particular update type. As"
Print "an example, assume this scenario. Note that the dates I am using are entirely made up..."
Print
Print "It is now Jan 2024. The latest version of Windows was released in Sep 2023. You download all the updates for Jan 2024."
Print "You notice that no Safe OS dynamic update has been released in Jan. You should then go back back in time to Dec 2023,"
Print "Nov 2023, etc. to find the most recent Safe OS Dynamic update. It is the most recent that you will want to use. Note"
Print "that it is possible that there may be no update of this type at all, especially early in the life of a new version."
Pause
GoTo HelpInjectUpdates

' Help Topic: Inject drivers into one or more Windows editions and create a multi edition bootable image

HelpInjectDrivers:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Inject Windows drivers into one or more Windows editions";
Locate 4, 38
Print "                and create a multi edition bootable image               ";
Locate 9, 1
Color 15
Print "    1) General information about this routine"
Print "    2) Acceptable Windows drivers"
Print "    3) How to obtain drivers"
Print "    4) Working with dual architecture images (applies only to dual architecture edition of this program)"
Print "    5) Organizing drivers"
Print
Color 0, 13
Print "    6) Return to main help menu "
Locate 28, 0
Color 15
Input "   Please select the item you would like help with by entering its number (6 returns to the main help menu): ", MenuSelection

Select Case MenuSelection
    Case 1
        GoTo InjectDriversGeneralInfo
    Case 2
        GoTo AcceptableDrivers
    Case 3
        GoTo ObtainingDrivers
    Case 4
        GoTo WorkWithDualArcImages_2
    Case 5
        GoTo OrganizingDrivers
    Case 6
        GoTo ProgramHelp
End Select

' We arrive here if the user makes an invalid selection from the menu

Color 15
Cls
Color 14, 4
Print
Print "You have made an invalid selection.";
Color 15
Print " You need to make a selection by entering a number from 1 to 6."
Pause
GoTo HelpInjectDrivers

' Help Topic: Inject drivers into one or more Windows editions and create a multi edition bootable image > General information about this routine

InjectDriversGeneralInfo:

Cls
Print "General Information About This Routine"
Print "======================================"
Print
Print "The routine to inject drivers is very similar to injecting Windows updates, however, in this routine, when specifying"
Print "the location of the drivers, the program will search all subdirectories for drivers."
Print
Print "As with the routine to inject Windows updates, this routine requires Windows ISO images with an install.wim and not an"
Print "INSTALL.ESD. Please see the help for the routine that injects Windows updates for a more detailed discussion of"
Print "acceptable Windows images."
Print
Print "Please note that drivers are injected into only the Windows editions whose indices you specify. For example, if you have"
Print "a Windows image with 11 editions, and you add drivers to index 6, only the edition of Windows at index 6 will have those"
Print "drivers. This routine will allow you to add drivers to multiple, or even all, indices. However, you can also run this"
Print "routine more than once in order to install different drivers into different editions of Windows."
Pause
GoTo HelpInjectDrivers

' Help Topic: Inject drivers into one or more Windows editions and create a multi edition bootable image > Acceptable Windows drivers

AcceptableDrivers:

Cls
Print "Acceptable Windows Drivers"
Print "=========================="
Print
Print "Drivers need to have the .INF file(s) accessible. If drivers are packaged in .CAB files, this program has a routine"
Print "available to allow files to be extracted from the .CAB files. Run that routine first to extract the drivers."
Print
Print "This process will NOT inject Windows updates. This routine only injects drivers. There is a separate option from the"
Print "main menu that will allow Windows updates to be injected."
Pause
GoTo HelpInjectDrivers

' Help Topic: Inject drivers into one or more Windows editions and create a multi edition bootable image > How to obtain drivers

ObtainingDrivers:

Cls
Print "How to Obtain Drivers"
Print "====================="
Print
Print "You can download drivers from the Microsoft update catalog, the manufacturers web site, or elsewhere, so long as you"
Print "can extract the drivers so that the .INF files are accessible."
Pause
GoTo HelpInjectDrivers

' Help Topic: Inject drivers into one or more Windows editions and create a multi edition bootable image > Working with dual architecture images (applies only to dual architecture edition of this program

WorkWithDualArcImages_2:

Cls
Print "Working with Dual Architecture Images (Applies Only to Dual Architecture Edition of This Program)"
Print "================================================================================================="
Print
Print "The same rules for working with dual architecture images apply when injecting drivers as for injecting Windows updates."
Print "Please see this topic in the help section for injecting Windows updates."
Pause
GoTo HelpInjectDrivers

' Help Topic: Inject drivers into one or more Windows editions and create a multi edition bootable image > Organizing your update files

OrganizingDrivers:

Cls
Print "Organizing Drivers"
Print "=================="
Print
Print "Organize your drivers like this:"
Print
Print "D:\Drivers"
Print "      \x64"
Print "            \Desktop Computer"
Print "            \Laptop"
Print "            \Server"
Print "      \x86"
Print "            \Tablet"
Print
Print "If you want to inject the drivers for one computer, when the computer asks for the location of the drivers, provide the"
Print "name of that folder. Example: D:\Drivers\Laptop. If you want to inject the drivers for all systems, specify D:\Drivers."
Print "This works because the program will install all drivers in the folder that you specify as well as all sub folders."
Print
Print "If you are working only with x64 Windows editions then no x86 folder is needed."
Pause
GoTo HelpInjectDrivers

' Help Topic: Inject boot-critical drivers into one or more Windows editions and create a multi edition bootable image

HelpInjectBCD:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Inject boot-critical drivers into one or more Windows ";
Locate 4, 38
Print "                editions and create a multi edition bootable image    ";
Locate 9, 1
Color 15
Print "This routine operates in the same manner as the routine to inject drivers. Please see the help for that routine for"
Print "details. However, please note that since the boot.wim (Windows PE) is shared between all editions of Windows in the"
Print "image, any boot critical drivers that you inject using this routine will be available to all editions of Windows in"
Print "your image."
Pause
GoTo ProgramHelp

' Help Topic: Inject registry entries into BOOT.WIM to bypass Windows 11 requirements

HelpBypassWin11Requirements:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " Dual Architecture Edition         "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print "Program Help - Modify Windows ISO to bypass system requirements and "
Locate 4, 38
Print "               optionally force use of previous version of setup    ";
Locate 9, 1
Color 15
Print "Bypassing Windows 11 System Requirements"
Print
Print "This routine will make modifications to bypass the check for Windows 11 system requirements. This affects both"
Print "clean and upgrade installs of Windows."
Print
Print "Please note that the requirement for a CPU to support SSE4.2 instructions still remains. That cannot be changed."
Print
Print "A Windows image with these modifications can still be used with a system that meets the Windows 11 requirements."
Print "In that instance, it simply prevents the check for those requirements."
Print
Print "Forcing the Use of the Previous Version of Setup"
Print
Print "Windows 11 24H2 introduces a new setup experience. When installing Windows manually (not unattended), the setup GUI"
Print "will give you the opportunity to run the previous version of setup. For unattended setup, there is no such option so"
Print "this option will force the previous version of setup to be used. It will will also force the previous version to be"
Print "used for manual installation, so this can be helpful if you wish to make this the default. Note that there will not"
Print "be an option to use the new setup experience."
Pause
GoTo ProgramHelp

' Help Topic: Make or update a bootable drive from one or more Windows / WinPE / WinRE ISO images

HelpMakeMultiBootImage:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Make or update a bootable drive from one or ";
Locate 4, 38
Print "                more Windows / WinPE / WinRE ISO images     ";
Locate 9, 1
Color 15
Print "    1) General information about this routine"
Print "    2) Disk limitations"
Print
Color 0, 13
Print "    3) Return to main help menu "
Locate 28, 0
Color 15
Input "   Please select the item you would like help with by entering its number (3 returns to the main help menu): ", MenuSelection

Select Case MenuSelection
    Case 1
        GoTo MakeBootDriveHelp
    Case 2
        GoTo DiskLimitationsHelp
    Case 3
        GoTo ProgramHelp
End Select

' We arrive here if the user makes an invalid selection from the menu

Color 15
Cls
Color 14, 4
Print
Print "You have made an invalid selection.";
Color 15
Print " You need to make a selection by entering a number from 1 to 3."
Pause
GoTo HelpMakeMultiBootImage

' Help Topic: Make or update a bootable drive from one or more Windows / WinPE / WinRE ISO images
' > General information about this routine

MakeBootDriveHelp:

Cls
Print "General Information About This Routine"
Print "======================================"
Print
Print "There are two options available:"
Print
Print "1) Create or refresh standard boot media created from a single Windows ISO image"
Print "2) Create or refresh boot media to allow booting from an unlimited number of ISO images"
Print
Print "Common to Both Options"
Print "======================"
Print
Print "Both of the available options have these things in common."
Print
Print "Both options will create a bootable disk (USB Flash Disk, HDD, SSD, etc.) that is universal in nature. What this means"
Print "is that the media can be booted on both x64 and x86 systems, and both BIOS and UEFI based systems. This program can be"
Print "used to make bootable media using x64, x86, or dual architecture images and the media can use either an install.esd or"
Print "an install.wim image file. In order to accomplish this level of compatibility, we make use of two partitions on the"
Print "media, and you can create up to two additional partitions that can be used for anything you desire. The program will"
Print "also offer to automatically BitLocker encrypt any of the additional partitions that you create."
Pause
Cls
Color 0, 10: Print "Option 1:";: Color 15: Print " Create or refresh standard boot media created from a single Windows ISO image"
Print
Print "This option will create a standard Windows bootable drive. The media created by this program is immediately available"
Print "to be booted. An option to create additional partitions on the bootble media is available that will allow you to store"
Print "additional data on the same disk so that remaining space on the disk is not wasted."
Print
Color 0, 10: Print "Option 2:";: Color 15: Print " Create or refresh boot media to allow booting from an unlimited number of ISO images"
Print
Print "This option will create a customized Windows PE installation. After the disk is created, you must copy all the ISO image"
Print "files that you wish to boot from to the ";: Color 0, 14: Print "ISO Images";: Color 15: Print " folder on the second volume. These can include Windows installation"
Print "images, Windows recovery disks, or WinPE / WinRE based images such as a Macrium Reflect boot disk."
Print
Print "As with option 1, you can create additional partitions on your media if you wish."
Print
Print "A folder called ";: Color 0, 14: Print "Answer Files";: Color 15: Print " is also created on the 2nd volume to allow you to keep any unattended answer files. You can"
Print "name the answer files anything that you wish, just make certain to use a ";: Color 0, 14: Print ".XML";: Color 15: Print " file extension. When you select an ISO"
Print "image to boot, you will also be asked if you want to use an answer file if there are any available in this folder. The"
Print "answer file that you select will be copied to volume 1 and renamed to ";: Color 0, 14: Print "autounattend.xml";: Color 15: Print "."
Print
Print "NOTE: If you do not select an answer file from the list then you will be asked if you wish to create an answer file"
Print "on the fly. This will ask a series of question and your responses are used to generate an answer file on the fly."
Print
Print "Be careful when using an answer file. Depending upon how the answer file is configured, when you boot a disk with an"
Print "answer file present, it can wipe a disk without any warning!"
Print
Print "Finally, a folder called ";: Color 0, 14: Print "Other";: Color 15: Print " is created on the 2nd volume to allow you to keep any other files you may need."
Pause
Cls
Print "Preparing the Disk After Creating it"
Print "===================================="
Print
Print "After you have made the disk with this program, you will need to configure it to boot the ISO image of your choice."
Print "There are two ways to do this."
Print
Print "   1) Configure the Disk Within Windows"
Print
Print "Simply run the file named ";: Color 0, 14: Print "Config_UFD.bat";: Color 15: Print ". This will ask you what ISO image you wish to make bootable and it will allow"
Print "you to optionally select an answer file to use if any are present. Booting from the disk will then boot the selected ISO"
Print "image along with your selected answer file, if you selected one."
Print
Print "   2) Configure the Disk by Booting From it"
Print
Print "If you have not already configured the disk to boot a specific ISO image, then booting from it will cause the"
Color 0, 14: Print "Config_UFD.bat";: Color 15: Print " file to be run automatically, allowing you to configure the ISO image and answer file to boot with. This"
Print "is especially useful if you have a system that you cannot boot into Windows. When you then reboot and boot from this"
Print "disk once again, it will boot the ISO image using any answer file that you may have selected."
Print
Print "Checking the Disk Status and Resetting"
Print "======================================"
Print
Print "Once you have configured the disk to boot an ISO image, if you run the ";: Color 0, 14: Print "Config_UFD.bat";: Color 15: Print " file on the 2nd volume, it will"
Print "report to you which ISO image will be booted. It will also give you the option to reset the disk back to the original"
Print "configuration. If you choose to perform this reset, it will revert the disk back to the state that it was in after you"
Print "first created it, ready for you to select another image to make bootable. Please note that any ISO images, answer"
Print "files, and other files that you have placed on the disk will not be affected and will remain undisturbed."
Pause
Cls
Print "The Wipe and Refresh Selections"
Print "==============================="
Print
Print "For both the option to create a standard bootable disk or to create a disk allowing boot from multiple ISO images, you"
Print "will be given a choice to Wipe or Refresh the disk. The first time you run this program, you need to select the Wipe"
Print "option. This will completely wipe out the disk that you are creating, partitioning and formatting it for first time use."
Print "After the initial creation, you can choose the Refresh option to recreate the disk while still leaving additional data"
Print "and partitions intact. In the case where you have created a disk with a single Windows boot image, the first two"
Print "partitions will be recreated using whatever Windows image you wish. It can be a totally different image than you used"
Print "originally. Any additional partitions that you created will be left undisturbed. In other words, if you simply want to"
Print "change the version of Windows that this disk will boot, or if you have updated your image in any way, simply perform a"
Print "refresh to update the image without disturbing data on other partitions."
Print
Print "If you created a disk allowing you to choose from multiple ISO images to boot, then a refresh operation will update the"
Print "version of Windows PE to the version currently on your system. All the elements needed to make the disk properly boot"
Print "will be recreated, but the ";: Color 0, 14: Print "Answer Files";: Color 15: Print ", ";: Color 0, 14: Print "ISO Images";: Color 15: Print ", and ";: Color 0, 14: Print "Other";: Color 15: Print " folders on the second volume will be left intact. In"
Print "addition, any other partitions on the disk will be left alone. Be aware that for a disk created with this option, you"
Print "should run the ";: Color 0, 14: Print "Config_UFD.bat";: Color 15: Print " on the 2nd partition to revert back to the original state before you can perform a refresh"
Print "operation. If your goal is to update the images themselves, then simply copy the updated ISO image files to the ISO"
Print "images folder on the 2nd partition."
Pause
Cls
Print "Unattended Answer File Generator"
Print "================================"
Print
Print "When you run config_ufd.bat on a disk that you created with the option to boot from multiple ISO images, one option"
Print "that is presented to you is the option to specify an unattended answer file"










GoTo HelpMakeMultiBootImage

' Help Topic: Make or update a bootable drive from one or more Windows / WinPE / WinRE ISO images
' > Disk limitations

DiskLimitationsHelp:

Cls
Print "Disk Limitations"
Print "================"
Print
Print "Be aware that for greatest compatibility, you should use media that is no larger than 2 TB in size. If you use media"
Print "that is larger than 2 TB in size, the program will give you the option to limit the media to 2 TB for the greatest "
Print "compatibility, or to initialize the disk to its full capacity, but sacrificing the ability to be booted on legacy"
Print "BIOS based systems."
Print
Print "You can use any rewritable media that your system is a able to boot from. UFD (USB Flash Disk), HDD, SD Card, are all"
Print "examples of valid media, assuming that your systems is capable of booting from that media."
Print
Print "Due to the frequently slow nature of some UFDs (USB Flash Drives), it is suggested that you use a UFD with a"
Print "reasonable level of performance. Option 2 will read from and write to the same disk when configuring the media"
Print "after selecting an image to make bootable, so faster media can make a big difference in how long this takes."
Pause

GoTo HelpMakeMultiBootImage

' Help Topic: Create a bootable Windows ISO image that can include multiple editions

HelpMakeBootDisk2:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Create a bootable Windows ISO image ";
Locate 4, 38
Print "                that can include multiple editions "
Locate 9, 1
Color 15
Print "This routine will take Windows editions from one or more ISO images and combine them into a single ISO image. This"
Print "routine simply combines existing Windows editions into one image, it does not inject Windows updates or drivers."
Pause
GoTo ProgramHelp

' Help Topic: Create a bootable ISO image from Windows files in a folder

HelpCreateBootableISOFromFiles:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Create a bootable ISO image from Windows files in a folder "
Locate 9, 1
Color 15
Print "This routine will take a Windows image that has been extracted to a hard disk and make a bootable Windows ISO image"
Print "from it. This is especially useful if you want to alter or manipulate any files in the image. Once all the files are"
Print "organized as desired on the hard disk, use this routine to recreate an ISO image."
Pause
GoTo ProgramHelp

' Help Topic: Reorganize the contents of a Windows ISO image

HelpChangeOrder:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Reorganize the contents of a Windows ISO image ";
Locate 9, 1
Color 15
Print "If an image with multiple Windows editions needs to be reorganized, for example, in order to change the order in which"
Print "the different Windows editions are displayed in the boot menu, use this routine to do so. In addition to reordering"
Print "the Windows editions, editions of Windows can be entirely removed from the image."
Pause
GoTo ProgramHelp

' Help Topic: Convert between an ESD and WIM either standalone or in an ISO image

HelpConvertEsd:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " Dual Architecture Edition         "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Convert between an ESD and WIM either standalone or in an ISO image ";
Locate 9, 1
Color 15
Print "Most routines in this program require Windows images that contain an install.wim file. However, it's possible that"
Print "you may only have an image with an install.esd file. These files are smaller, so they are a good choice for"
Print "distribution over the internet, but they cannot be used to service the images."
Print
Print "This routine will convert an image with an install.wim file into an image with an install.esd file. You can then"
Print "use that image for servicing (injecting updates, drivers, etc.)."
Print
Print "This routine also can convert in the other direction, converting a WIM into an ESD image."
Print
Print "The conversion is automatic; if the original ISO image contains an install.esd it will be converted into an"
Print "install.wim. Likewise, an image with an install.wim will automatically be converted to one with an install.wim."
Print
Print "This routine can also convert between a standalone ESD and WIM image that is not inside of an ISO image."
Print
Print "Since this is the x64 only edition of the program, we support only conversion of single architecture Windows images"
Print "and not dual architecture images."
Pause
GoTo ProgramHelp

' Help Topic: Get image info - display basic info for each edition in an ISO image and display Windows build number

HelpGetWimInfo:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Get image info - display basic info for each edition in an ISO ";
Locate 4, 38
Print "                image and display Windows build number                         ";
Locate 9, 1
Color 15
Print "There are times where it may be necessary to know what index number is associated with a particular Windows edition,"
Print "or how many editions are stored in an image, or to view the NAME and DESCRIPTION metadata for Windows editions. This"
Print "routine will display that information and optionally save the output to a text file."
Print
Print "This routine will also display the build number of the Windows editions in the image. Note that only one edition of"
Print "Windows is actually checked for this information and that we assume all other editions are the same build."
Print
Print "In addition, we provide the option to display the build number of both the boot.wim and winre.wim. Please note we"
Print "can display the build number for the winre.wim only if your ISO image uses an install.wim file, not an install.esd."
Pause
GoTo ProgramHelp

' Help Topic: Modify the NAME and DESCRIPTION values for entries in a WIM file

HelpNameAndDescription:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Modify the NAME and DESCRIPTION values for entries in a WIM file ";
Locate 9, 1
Color 15
Print "When creating a Windows image with multiple editions of Windows, each edition must have valid NAME and DESCRIPTION"
Print "metadata. This information is displayed in the boot menu when booting from the image or media created from that image."
Print "When booting from the image or media, a list of operating systems is shown. This names in the list show the NAME"
Print "associated with the edition of Windows. As you scroll through that list, as each entry is selected, below the window"
Print "with the list of operating systems is a description of the currently selected operating system. This description is"
Print "the DESCRIPTION metadata."
Print
Print "Note that ISO images from Microsoft, the NAME and DESCRIPTION are often the same. However, you can set this to anything"
Print "you wish. Example:"
Print
Print "NAME: Windows 10 Pro"
Print "DESCRIPTION: Includes all drivers for HP laptop"
Pause
GoTo ProgramHelp

' Help Topic: Export drivers from this system

HelpExportDrivers:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Export drivers from this system ";
Locate 9, 1
Color 15
Print "This routine will export all drivers from the system on which it is run. Those drivers can then be injected into a"
Print "Windows image. In addition, this routine will create a batch file named Install_Drivers.bat in the folder with the"
Print "drivers. If a clean install of Windows is ever performed on that system, all the drivers can be restored by simply"
Print "running that one batch file."
Pause
GoTo ProgramHelp

' Help Topic: Expand drivers supplied in a .CAB file

HelpExpandDrivers:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Expand drivers supplied in a .CAB file ";
Locate 9, 1
Color 15
Print "In order to inject drivers into a Windows image, the .INF file(s) need to be accessible. This routine will expand"
Print "drivers supplied in a .CAB file so that the .INF file(s) are accessible."
Pause
GoTo ProgramHelp

' Help Topic: Create a Virtual Disk (VHDX)

HelpCreateVHDX:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Create a Virtual Disk (VHDX) ";
Locate 9, 1
Color 15
Print "This routine will create a Virtual Disk. This can be helpful in several different scenarios. For routines that won't"
Print "allow the use of a removable drive, a VHD can be created on the removable drive as a workaround. It can also be used"
Print "as temporary storage for data, etc."
Print
Print "Please note that starting with Windows 11 23H2, Windows has a new GUI for performing this operation directly and"
Print "easily from the OS. Simply search Windows Settings for the term "; Chr$(34); "Virtual Disk"; Chr$(34); "."
Pause
GoTo ProgramHelp

' Help Topic: Create a VHD, deploy Windows to it, and add it to the boot menu to make a dual boot configuration

HelpAddVHDtoBootMenu:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Create a VHD, deploy Windows to it, and add it to ";
Locate 4, 38
Print "                the boot menu to make a dual boot configuration   ";
Locate 9, 1
Color 15
Print "This routine will create a new installation of Windows on a VHD that can be booted on a physical machine in addition to"
Print "your already existing Windows installation without the need for any virtualization software."
Print
Print "You can deploy Windows to any partition, on an internal or USB connected HDD, SSD, or flash drive but the drive should"
Print "not be BitLocker encrypted. If C: is BitLocker encrypted, install to a drive or partition other than C:."
Print
Print "Please be aware that because this program updates the BCD store to add the new installation of Windows to the boot menu,"
Print "we need to temporarily suspend BitLocker if your C: drive is encrypted. We will do this automatically and BitLocker will"
Print "automatically be re-enabled the next time you boot into Windows on the C: drive."
Print
Print "Make certain to set the time in your new Windows installation. Not setting it can alter the real time clock causing your"
Print "primary installation of Windows to show the wrong time."
Print
Print "If you wish to delete the newly deployed copy of Windows, first suspend BitLocker on the C: drive (if enabled) either"
Print "via the GUI or run the command "; Chr$(34); "manage-bde -protectors -disable C:"; Chr$(34); ". This will suspend BitLocker until the system is"
Print "rebooted at which time BitLocker will automatically be re-enabled. Then run MSCONFIG, go to the boot tab, set your"
Print "primary Windows installation to the default, and delete the entry for the installation of Windows deployed to the VHD."
Print "You can then optionally delete the VHD to which you deployed Windows."
Pause
GoTo ProgramHelp

' Help Topic: Create a generic ISO image and inject files and folders into it

HelpCreateISOImage:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Create a generic ISO image and inject files and folders into it ";
Locate 9, 1
Color 15
Print "This routine takes the contents of any folder, including all subdirectories, and creates an ISO image from them. An ISO"
Print "image can be handy because it can often be created much quicker than a ZIP or other compressed file format. It can also"
Print "be easily attached to a VM to get files into a VM without the need for any network connectivity. Another good use for"
Print "an ISO image is to simply store an answer file for unattended installation. This can then be mounted to a VM to allow"
Print "fully unattended installation since Windows setup will search the root of all drives for an answer file."
Pause
GoTo ProgramHelp

' Help Topic: Cleanup files and folders

HelpGetFolderToClean:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Cleanup files and folders ";
Locate 9, 1
Color 15
Print "Routines that inject Windows updates and drivers into Windows images will mount an image using the DISM utility. If an"
Print "operation fails or is interrupted, it can leave stale mounted images in place. These mounts can persist even after a"
Print "reboot and will prevent you from having full access to those folders. This routine will attempt to cleanup those files"
Print "and folders."
Pause
GoTo ProgramHelp

' Help Topic: Unattended answer file generator

HelpAnswerFileGen:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Unattended answer file generator ";
Locate 9, 1
Color 15
Print "This routine will generate an autounattend.xml answer file for unattended setup of Windows. It covers mainly just the"
Print "basic settings needed for fully unattended installation. The idea is to keep this simple and reliable."
Print
Print "Current features:"
Print
Print " - Can create an answer file for either BIOS or UEFI systems"
Print " - For UEFI systems, allow a user defind Windows partition size and assign remaining space to another partition"
Print " - Allow bypassing Win 11 system requirements"
Print " - Allow bypassing installation of quality updates during setup"
Print " - Allow bypassing of automatic device encryption"
Print " - Complete user control of all partition sizes on Windows drive during setup"
Print
Print "Further enhancements to this feature are planned for the future."
Pause
GoTo ProgramHelp

' Help Topic: Exit

HelpExit:

Color 15
Cls
Color 0, 9
Print
Print " Windows Image Manager (WIM) Tools "
Print " x64 Only Edition                  "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 0, 10
Locate 3, 38
Print " Program Help - Exit ";
Locate 9, 1
Color 15
Print "Take a wild guess what this does."
Pause
GoTo ProgramHelp


' ********
' * Exit *
' ********

ProgramEnd:

Cls
Print "Performing a cleanup and exiting the program..."

' Remove the AV exclusion to the current program name

Cmd$ = "powershell.exe -command Remove-MpPreference -ExclusionProcess " + "'" + Chr$(34) + Command$(0) + Chr$(34) + "'"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

System


' *******************************************************************************************
' * Sub PROCEDURES go here. These are subroutines that are still a part of the main program *
' * and that are not in SUB / END SUB blocks.                                               *
' *******************************************************************************************


DisplayIndices2:

' Display a list of Windows editions and associated indices in the selected image file
' Note that this routine creates a file called Image_Info.txt. If you don't need it after
' a return from this subroutine, make sure to delete it.
'
' NOTE: Before calling the subroutine, set Silent$ to "Y" or "N". If set to "Y" then the routine will
' run silently. In other words, it save results to the Image_Info.txt file but will not display it.

If Silent$ = "N" Then
    Cls
    Print "Preparing to display a list of indices...."
    Print
Else
    Print
    Print "Building a list of available editions."
End If

MountISO SourcePath$

Shell "echo File Name: " + Mid$(SourcePath$, _InStrRev(SourcePath$, "\") + 1) + " > Image_Info.txt"
Shell "echo. >> Image_Info.txt"
Shell "echo ***************************************************************************************************** >> Image_Info.txt"
Shell "echo * Below is the list of Windows editions and associated indicies available for the above named file. * >> Image_Info.txt"
Shell "echo ***************************************************************************************************** >> Image_Info.txt"
Shell "echo. >> Image_Info.txt"

' The lines below test to see if this image has an install.esd or an install.wim and runs the appropriate command.
' Normally, we should not need this. Only an install.wim should be present for this project, but this routine can handle either.

InstallFileTest$ = MountedImageDriveLetter$ + ArchitectureChoice$ + "\sources\install.wim"

If _FileExists(InstallFileTest$) Then
    InstallFile$ = "\sources\install.wim >> Image_Info.txt"
Else
    InstallFile$ = "\sources\install.esd >> Image_Info.txt"
End If

Cmd$ = "dism /Get-WimInfo /WimFile:" + MountedImageDriveLetter$ + ArchitectureChoice$ + InstallFile$
Shell Cmd$

' Dismount the ISO image

Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourcePath$ + "'" + Chr$(34) + Chr$(34)
Shell _Hide Cmd$

' Display Image_Info.txt which lists all of the indicies

If Silent$ = "N" Then
    Cls
    DisplayFile "Image_Info.txt"
End If

Return


' End of Program

EndProgram:

System

End

' ****************************************
' * DATA strings will be specified here. *
' ****************************************

DriveLetterData:

Data C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z


' ********************************************************************************
' * Below are subroutines that need to appear after the end of the program.      *
' * Subroutines called by name reference and placed in SUB / END SUB blocks must *
' * be placed after the end of the main program.                                 *
' ********************************************************************************


Sub CleanPath (Path As String)

    Dim Path1 As String

    Path1$ = Path$

    ' Remove quotes and trailing backslash from a path

    ' To use this subroutine: Pass the path to this sub, the sub will return the path
    ' without a trailing backslash in Temp$.

    Dim x As Integer

    ' start by stripping the quotes

    Temp$ = ""

    For x = 1 To Len(Path1$)
        If Mid$(Path1$, x, 1) <> Chr$(34) Then
            Temp$ = Temp$ + Mid$(Path1$, x, 1)
        End If
    Next x

    ' Remove the trailing backslash, if present

    If Right$(Temp$, 1) = "\" Then
        Temp$ = Left$(Temp$, (Len(Temp$) - 1))
    End If

End Sub


Sub FileTypeSearch (Path$, FileType$, SearchSubFolders$)

    ' This routine will receive a path and a file name extension. It will search the path specified
    ' for any occurances of files with the specified extension. It will return the number of those files
    ' found in NumberOfFiles and each of those file names in an array called TempArray$().

    ' The path passed to this subroutine should end with a trailing backslash. Example: D:\MyFolder\
    ' Pass the extension with a period. Example: .ISO. For all files regardless of extension, pass
    ' a "*" with no period in front of it.
    '
    ' Pass a "N" if you do not want to search subfolders, or a "Y" if you do.


    ' Initialize the variables

    Dim Path_Local As String
    Dim FileType_Local As String
    Dim SearchSubFolders_Local As String
    Dim Cmd As String
    Dim file As String

    Path_Local$ = Path$
    FileType_Local$ = FileType$
    SearchSubFolders_Local = SearchSubFolders$

    NumberOfFiles = 0 ' Set initial value
    FileType_Local$ = UCase$(FileType_Local$)

    ' Build the command to be run

    If FileType_Local$ = "*" Then
        Select Case SearchSubFolders_Local$
            Case "N"
                Cmd$ = "DIR /B " + Chr$(34) + Path_Local$ + Chr$(34) + " > WIM_TEMP.TXT"
            Case "Y"
                Cmd$ = "DIR /B " + Chr$(34) + Path_Local$ + Chr$(34) + " /s" + " > WIM_TEMP.TXT"
        End Select
    Else
        Select Case SearchSubFolders$
            Case "N"
                Cmd$ = "DIR /B " + Chr$(34) + Path_Local$ + "*" + FileType_Local$ + Chr$(34) + " > WIM_TEMP.TXT"
            Case "Y"
                Cmd$ = "DIR /B " + Chr$(34) + Path_Local$ + "*" + FileType_Local$ + Chr$(34) + " /s" + " > WIM_TEMP.TXT"
        End Select
    End If

    Shell _Hide Cmd$

    If _FileExists("WIM_TEMP.TXT") Then
        Open "WIM_TEMP.TXT" For Input As #1
        Do Until EOF(1)
            Line Input #1, file$
            If FileType_Local$ = "*" Then
                If file$ <> "File Not Found" Then
                    NumberOfFiles = NumberOfFiles + 1

                    If Left$(file$, 1) = "-" Then
                        TempArray$(NumberOfFiles) = "-" + Path_Local$ + Right$(file$, (Len(file$) - 1))
                    Else
                        TempArray$(NumberOfFiles) = Path_Local$ + file$
                    End If
                End If
            ElseIf UCase$(Right$(file$, 4)) = UCase$(FileType_Local$) Then
                NumberOfFiles = NumberOfFiles + 1

                ' In case we are injecting drivers, we would be searching for ".INF" files here. For these files, we have no reason to store the name of these files
                ' because we don't process the files one by one. All we need for these is confirmation that .INF files exist.

                If FileType_Local$ <> ".INF" Then
                    TempArray$(NumberOfFiles) = Path_Local$ + file$
                End If

            End If
        Loop
        Close #1
        Kill "WIM_TEMP.TXT"
    End If

End Sub


Sub MountISO (ImagePath$)

    ' This routine will mount the ISO image at the path passed from the main program and
    ' will get the CDROM ID (Ex. \\.\CDROM0) and save in MountedImageCDROMID$. It will also
    ' get the drive letter from and save in MountedImageDriveLetter$ (Ex. E:).

    Dim ImagePath_Local As String
    Dim Cmd As String
    Dim GetLine As String
    Dim count As Integer

    ImagePath_Local$ = ImagePath$
    MountedImageCDROMID$ = ""
    MountedImageDriveLetter$ = ""

    CleanPath (ImagePath_Local$)
    ImagePath_Local$ = Temp$
    Cmd$ = "powershell.exe -command " + Chr$(34) + "Mount-DiskImage " + Chr$(34) + "'" + ImagePath_Local$ + "'" + Chr$(34) + Chr$(34) + " > MountInfo1.txt"
    Shell Cmd$
    Cmd$ = "powershell.exe -command " + Chr$(34) + "Get-DiskImage -ImagePath '" + ImagePath_Local$ + "' | Get-Volume" + Chr$(34) + " > MountInfo2.txt"
    Shell Cmd$
    Open "MountInfo1.txt" For Input As #1
    Open "MountInfo2.txt" For Input As #2

    Do Until EOF(1)
        Line Input #1, GetLine$
        If InStr(1, GetLine$, "\\.\CDROM") Then
            MountedImageCDROMID$ = Right$(GetLine$, Len(GetLine$) - ((InStr(1, GetLine$, "\\.\CDROM"))) + 1)
            Exit Do
        End If
    Loop

    Close #1

    For count = 1 To 4
        If EOF(2) Then

            ' We ran into a problem with this file. We were expecting at least 4 lines of text but did not find it.

            Exit For
        End If

        Line Input #2, GetLine$
    Next count

    MountedImageDriveLetter$ = Left$(GetLine$, 1) + ":"
    Close #2
    Kill "MountInfo1.txt"
    Kill "MountInfo2.txt"

    If ((MountedImageCDROMID$ = "") Or (MountedImageDriveLetter$ = "")) Then
        Cls
        Color 14, 4: Print "WARNING!";: Color 15: Print " While trying to determine the ID and drive letter of a mounted image we encountered a fatal error."
        Print "The program will terminate."
        Print
        Print "Diagnostic info:"
        Print
        Print "The CDROM ID returned was: "; MountedImageCDROMID$
        Print "The CDROM drive letter returned was: "; MountedImageDriveLetter$
        Pause
        System
    End If

End Sub


Sub Cleanup (CleanupPath$)

    ' Pass the name of the folder to cleanup to this routine. It will return
    ' CleanupSuccess=0 if it failed, CleanupSuccess=1 if successful, CleanupSuccess=2 if
    ' the specified folder does not exist. NOTE: Currently, nothing is checking for a
    ' return of 2. This routine itself simply displays a message if the folder does not exist.

    ' NOTE: This routine wants a path with a "\" at the end.

    ' When injecting either Windows updates or drivers into an ISO image, if
    ' the process is aborted without being allowed to finish, files may be
    ' left behind that you cannot delete manually. This routine will try to
    ' correct that situation. Note that this routine will only try to delete folders named
    ' "Mount", "ISO_Files", and "Scratch". This assures that we maintain any log files in
    ' the "LOGS" folder.

    ' If the files still exist after the intial cleanup attempt,  then we will attempt
    ' a fix by closing any open DISM session. If that still fails, we will inform the
    ' user that a reboot may be needed.

    ' After an automatic cleanup attempt, we will set the variable AutoCleanup to 1 so that we know that a
    ' cleanup has already been attempted. That way, if we fail again we know that we need to abort and
    ' warn the user.

    ' Initialize AutoCleanup to 0 before the start of the cleanup process.

    ' First, let's check to see if the specified folder even exists. If not, then there
    ' is nothing to cleanup but we still consider that successful because there are no
    ' unwanted files present.

    Dim AutoCleanup As Integer
    Dim TempPath As String
    Dim Cmd As String
    Dim Count As Integer
    Dim MountDir As String

    If Not (_DirExists(CleanupPath$)) Then
        Cls
        Color 14, 4: Print "There is nothing to cleanup because that folder does not exist!": Color 15
        Pause
        CleanupSuccess = 2
        GoTo NoSuchFolder
    End If

    AutoCleanup = 0 ' Set initial value

    Cls
    Print "Attempting cleanup of the folder: "
    Print
    Color 10: Print CleanupPath$: Color 15
    Print
    Print "Please standby. This may take a while, especially if there are many files or stale mounts..."

    StartCleanup:

    ' Cleaning up previous stale DISM operations

    Count = 0
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /get-mountedimageinfo > DismInfo.txt"
    Shell Chr$(34) + Cmd$ + Chr$(34)
    Open "DismInfo.txt" For Input As #1

    Do
        Count = Count + 1
        Line Input #1, (MountDir$)
        If (InStr(MountDir$, "Mount Dir :")) Then
            If Count = 1 Then
                Print
                Color 14, 4: Print "Warning!";: Color 15: Print " There is still at least one image mounted by DISM open."
                Print "We will try to clear these at this time. Please standby..."
            End If
            MountDir$ = Right$(MountDir$, (Len(MountDir$) - 12))
            Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /MountDir:" + Chr$(34) + MountDir$ + Chr$(34) + " /Discard"
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
            Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Cleanup-WIM"
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
            Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /cleanup-mountpoints"
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If
    Loop Until EOF(1)

    Close #1
    Kill "disminfo.txt"

    ' Run a check a second time for open mounts. The above procedure should have cleared any open mounts, but if we still
    ' have an open mount then we may need to reboot to resolve the situation.

    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /get-mountedimageinfo > DismInfo.txt"
    Shell Chr$(34) + Cmd$ + Chr$(34)
    Open "DismInfo.txt" For Input As #1

    Do
        Line Input #1, MountDir$
        If (InStr(MountDir$, "Mount Dir :")) Then
            Close #1
            Cls
            Print "We were not able to clear all mounts. Try rebooting the system and then run the program again."
            Pause
            Kill "disminfo.txt"
            System
        End If
    Loop Until EOF(1)

    Close #1
    Kill "disminfo.txt"

    Cmd$ = "c:\windows\system32\takeown /f " + Chr$(34) + CleanupPath$ + "*.*" + Chr$(34) + " /r /d y"
    Shell _Hide Cmd$
    Cmd$ = "icacls " + Chr$(34) + CleanupPath$ + "*.*" + Chr$(34) + " /T /grant %username%:F"
    Shell _Hide Cmd$

    ' Trying to cleanup the "SSU_x64" subfolder

    TempPath$ = CleanupPath$ + "SSU_x64\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "Mount" subfolder

    TempPath$ = CleanupPath$ + "Mount\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "ISO_Files" subfolder

    TempPath$ = CleanupPath$ + "ISO_Files\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "Scratch" subfolder

    TempPath$ = CleanupPath$ + "Scratch\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "Temp" subfolder

    TempPath$ = CleanupPath$ + "Temp\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "WIM_x86" subfolder

    TempPath$ = CleanupPath$ + "WIM_x86\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "WIM_x64" subfolder

    TempPath$ = CleanupPath$ + "WIM_x64\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "Assets" subfolder

    TempPath$ = CleanupPath$ + "Assets\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "WinRE_Mount" subfolder

    TempPath$ = CleanupPath$ + "WinRE_Mount\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "WinPE_Mount" subfolder

    TempPath$ = CleanupPath$ + "WinPE_Mount\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "WinPE" subfolder

    TempPath$ = CleanupPath$ + "WinPE\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "WinRE" subfolder

    TempPath$ = CleanupPath$ + "WinRE\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "Setup_DU_x64" subfolder

    TempPath$ = CleanupPath$ + "Setup_DU_x64\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' Trying to cleanup the "Setup_DU_x86" subfolder

    TempPath$ = CleanupPath$ + "Setup_DU_x86\"

    If _DirExists(TempPath$) Then
        Shell _Hide "rmdir /s /q " + Chr$(34) + TempPath$ + Chr$(34)
    End If

    If _DirExists(TempPath$) Then GoTo FoldersNotDeleted

    ' If we reach this point, then the folders we tried to delete were successfully deleted.

    GoTo DeletionSuccessful

    ' The following code handles the situation if we were not able to delete the folders.

    FoldersNotDeleted:

    ' If we arrive here, it means that trying to delete a folder failed.

    Cls
    Print "We were not able to delete one or more folders under this folder:"
    Print
    Color 10: Print CleanupPath$: Color 15
    Print
    Print "The program has already automatically attempted to correct the situation but was not able to"
    Print "do so. The most likely cause for this is that a previous run of this program may have been"
    Print "aborted while the Microsoft DISM utility still had files locked."
    Print
    Print "Please try rebooting the computer, and then try running this program again."
    Pause
    CleanupSuccess = 0
    GoTo EndCleanup

    DeletionSuccessful:

    CleanupSuccess = 1
    GoTo EndCleanup

    NoSuchFolder:

    CleanupSuccess = 2
    GoTo EndCleanup

    EndCleanup:

End Sub


Sub YesOrNo (YesNo$)

    ' This routine checks whether a user responded with a valid "yes" or "no" response. The routine will return a capital "Y" in YN$
    ' if the user response was a valid "yes" response, a capital "N" if it was a valid "no" response, or an "X" if not a valid response.
    ' Valid responses are the words "yes" or "no" or the letters "y" or "n" in any case (upper, lower, or mixed). Anything else is invalid.

    Select Case UCase$(YesNo$)
        Case "Y", "YES"
            YN$ = "Y"
        Case "N", "NO"
            YN$ = "N"
        Case Else
            YN$ = "X"
    End Select

End Sub


Sub DetermineArchitecture (SourcePath$, ChosenIndex)

    ' Pass a path, including the file name, of a Windows ISO image and an index number to this routine.
    ' It will return one of the following results in the variable called ImageArchitecture$: x64, x86, DUAL, NONE

    ' We need the index number without leading spaces so we are converting it to a string.

    Dim SourcePath_Local As String
    Dim ChosenIndex_Local As Integer
    Dim ChosenIndexString As String
    Dim Cmd As String
    Dim ReadLine As String
    Dim position As Integer

    SourcePath_Local$ = SourcePath$
    ChosenIndex_Local = ChosenIndex



    ChosenIndexString$ = Str$(ChosenIndex_Local)
    ChosenIndexString$ = Right$(ChosenIndexString$, ((Len(ChosenIndexString$) - 1)))

    ' Clear variable

    MountedImageDriveLetter$ = ""
    MountISO SourcePath_Local$

    If _FileExists(MountedImageDriveLetter$ + "\x64\sources\install.wim") Then
        ImageArchitecture$ = "DUAL"
        GoTo DetermineArchitectureExit
    End If

    If Not (_FileExists(MountedImageDriveLetter$ + "\sources\install.wim")) Then
        ImageArchitecture$ = "NONE"
        GoTo DetermineArchitectureExit
    End If

    Cmd$ = "dism /get-wiminfo /wimfile:" + MountedImageDriveLetter$ + "\sources\install.wim /index:" + ChosenIndexString$ + " > WIM_TEMP.TXT"
    Shell Cmd$
    Open "WIM_TEMP.TXT" For Input As #1

    Do
        Line Input #1, ReadLine$
        position = InStr(ReadLine$, "Architecture")
    Loop While position = 0

    Close #1
    Kill "WIM_TEMP.TXT"

    If Right$(ReadLine$, 3) = "x64" Then
        ImageArchitecture$ = "x64"
        GoTo DetermineArchitectureExit
    ElseIf Right$(ReadLine$, 3) = "x86" Then
        ImageArchitecture$ = "x86"
        GoTo DetermineArchitectureExit
    Else
        ImageArchitecture$ = "NONE"
        GoTo DetermineArchitectureExit
    End If

    DetermineArchitectureExit:

    Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourcePath_Local$ + "'" + Chr$(34) + Chr$(34)
    Shell _Hide Cmd$

End Sub


Sub RemovableDiskCheck (DiskLetterOrID$)

    ' This subroutine will determine whether a disk is removable or not.

    ' Pass a disk letter to this routine (C:, D:, etc.) or a disk ID (0, 1, 2, etc.) and this routine
    ' will check to see if it is removable or not. The variable IsRemovable will be set to "1" if
    ' the disk is removable and to "0" if it is not. It the disk letter or ID is not valid, a "2" will be returned.

    ' Determine if the value passed to this routine is a drive letter or a disk ID.
    ' If we passed a DISK ID (a number) then the value of DiskLetterOrID$ will be
    ' greater than 0 (0 is only used by the system disk). Otherwise, if a drive letter
    ' is passed, then the value of the string will be 0.
    '

    ' Local variable declarations

    ' ***************************************************************************
    ' * Make sure to define IsRemovable as a SHARED integer in the main program *
    ' ***************************************************************************

    Dim DiskID As Integer ' Used to hold the Disk ID number that was passed to this routine
    Dim LineCount As Integer ' Stores how many lines are contained in a file
    Dim MatchFound As Integer ' If user specifies a drive letter and we find that letter to be valid, we set MatchFound to "1" to indicate a match was found in diskpart
    Dim NumberOfLines As Integer ' The number of lines we have to read from file to get the information we need
    Dim Temp As String ' Temporary holding place for lines pf text read from file
    Dim x As Integer ' Loop counter for a FOR...NEXT loop
    Dim VolID As Integer ' When a drive letter is passed to this routine, we first determine what volume number it is located on and store that in VolID

    ' End local variable definitions

    IsRemovable = 2 ' Setting an intial value of "2" to indicate invalid letter or disk ID until we determine otherwise
    MatchFound = 0 ' Set initial value

    If Val(DiskLetterOrID$) > 0 Then
        DiskID = Val(DiskLetterOrID$)
        GoTo RemovableDiskCheck_DriveID
    Else
        GoTo RemovableDiskCheck_Letter
    End If

    RemovableDiskCheck_Letter:

    ' A drive letter was passed to the routine. Obtain the Volume ID on which this drive letter resides.

    Open "TEMP.BAT" For Output As #1
    Print #1, "@echo off"
    Print #1, "(echo list vol"
    Print #1, "echo exit"
    Print #1, ") | diskpart > DiskpartOut.txt"
    Close #1
    Shell "TEMP.BAT"
    If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

    ' Parse the DiskpartOut.txt file
    ' Start by determining how many lines there are

    LineCount = 0 ' Set initial value

    Open "DiskpartOut.txt" For Input As #1

    Do Until EOF(1)
        Line Input #1, Temp$
        LineCount = LineCount + 1
    Loop

    Close #1

    ' Lines 1 to 9 are header information that we want to skip past. After that, analyize each line
    ' to see if it holds the drive letter of interest.

    Open "DiskpartOut.txt" For Input As #1

    For x = 1 To 9
        Line Input #1, Temp$
    Next x

    For x = 10 To LineCount
        Line Input #1, Temp$
        If Len(Temp$) >= 16 Then
            If (Mid$(Temp$, 16, 1) = UCase$(Left$(DiskLetterOrID$, 1))) And (Mid$(Temp$, 3, 6) = "Volume") Then
                MatchFound = 1
            Else
                MatchFound = 0
            End If
        End If
        If MatchFound = 1 Then
            VolID = Val(Mid$(Temp$, 10, 2))
            Exit For
        End If
    Next x

    Close #1
    Kill "DiskpartOut.txt"

    ' If we have successfully retrieved the Volume ID for the letter that the user specified, we need determine what disk number (ID) this is located on.
    ' If not, then we need to report the disk selected as invalid.

    If MatchFound = 0 Then
        IsRemovable = 2
        GoTo RemovableDiskCheck_Done
    Else
        GoTo RemovableDiskCheck_GetDiskID
    End If

    ' Using the volume ID, we will now determine the disk ID on which the volume with the specified drive letter is located.

    RemovableDiskCheck_GetDiskID:

    Open "TEMP.BAT" For Output As #1
    Print #1, "@echo off"
    Print #1, "(echo sel vol"; VolID
    Print #1, "echo detail vol"
    Print #1, "echo exit"
    Print #1, ") | diskpart > DiskpartOut.txt"
    Close #1
    Shell "TEMP.BAT"
    If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

    ' Parse the DiskpartOut.txt file
    ' Lines 1 to 12 are header information that we want to skip past. We want info in line 13.

    Open "DiskpartOut.txt" For Input As #1

    For x = 1 To 13
        Line Input #1, Temp$
    Next x

    Close #1
    Kill "DiskpartOut.txt"

    ' As a simple check, make sure that characters 3 to 6  of the line selected have the word "Disk". If not,
    ' then something went wrong. Set IsRemovable to 2 and exit the routine.

    If Mid$(Temp$, 3, 4) <> "Disk" Then
        IsRemovable = 2
        GoTo RemovableDiskCheck_Done
    End If

    DiskID = Val(Mid$(Temp$, 8, 2))

    ' Start of section that checks disk ID
    ' If user passed to this routine a disk ID, we come directly here. Otherwise, if the specified a drive letter, then we will have already
    ' determined the disk ID , saved in the "DiskID" variable and dropped down to this routine.

    RemovableDiskCheck_DriveID:

    ' Check the drive ID to see if disk is removable

    Open "TEMP.BAT" For Output As #1
    Print #1, "@echo off"
    Print #1, "(echo select disk"; DiskID
    Print #1, "echo detail disk"
    Print #1, "echo exit"
    Print #1, ") | diskpart > DiskpartOut.txt"
    Close #1
    Shell "TEMP.BAT"
    If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

    ' Parse the DiskpartOut.txt file
    ' Start by determining how many lines there are

    LineCount = 0 ' Set initial value

    Open "DiskpartOut.txt" For Input As #1

    Do Until EOF(1)
        Line Input #1, Temp$
        LineCount = LineCount + 1
    Loop

    Close #1

    ' We are interested in the 4th line from the end so we are setting a variable called
    ' NumberOfLines to be LineCount - 3 and we will then analyze that line.

    Open "DiskpartOut.txt" For Input As #1
    NumberOfLines = LineCount - 3

    For x = 1 To NumberOfLines
        Line Input #1, Temp$
    Next x

    Close #1
    Kill "DiskpartOut.txt"

    ' If the disk is removable, then the 9 characters starting with the 39th character of Temp$ will equal "Removable".

    If (Len(Temp$) < 48) Then
        IsRemovable = 0
        GoTo RemovableDiskCheck_Done
    End If

    If Mid$(Temp$, 40, 9) = "Removable" Then
        IsRemovable = 1
    Else
        IsRemovable = 0
    End If

    ' End of section that checks the Disk ID to see if it is removable

    RemovableDiskCheck_Done:

End Sub


Sub GetWimInfo_Main (SourcePath$, GetWimInfo_Silent)

    ' This routine will save WIM info to a text file called Image_Info.txt located in the same
    ' folder that the program is run from. This routine can be run with status information or
    ' silently. Pass to this routine the name of the ISO image for which to get info in a
    ' string and a 0 or 1 to indicate silent mode (0 = run normally, 1 = run silent).
    '
    ' Note that if this routine is run silently, we do not ask for or gather information for
    ' the winre.wim and boot.wim images.
    '
    ' If you no longer need the file Image_Info.txt, make sure to delete it after a return from
    ' this routine.

    ' Declare local variables

    Dim A As String ' Hold contents of WIM_Info2.txt for parsing
    Dim Architecture As String ' Tracks the architecture type of the selected ISO image
    Dim Cmd As String ' Holds a string that has been built to be run with a "Shell" command
    Dim DestinationFolder As String ' Complete path to location where project can be created if boot.wim and winre.wim info is desired
    Dim DestinationIsRemovable As Integer
    Dim DriveLetter As String
    Dim ff As Integer ' Hold the next available free file number
    Dim InstallFile As String
    Dim InstallFileTest As String ' Used to test whether an install.wim file is present
    Dim InstallFileTest2 As String ' Used to test whether an install.esd file is present
    Dim LocalTemp As String ' Temporary data
    Dim SP_Build As String ' Second portion of the full Windows build number
    Dim ShowExtendedInfo As String ' Set to "Y" if info for boot.wim and winre.wim should be shown
    Dim Version As String ' First portion of the full Windows build number
    Dim X As Integer ' Temporary value used for manipulation of string
    Dim Y As Integer ' Temporary value used for manipulation of string

    ShowExtendedInfo$ = "N" ' Set initial value

    Shell "echo File Name: " + Mid$(SourcePath$, _InStrRev(SourcePath$, "\") + 1) + " > Image_Info.txt"
    Shell "echo. >> Image_Info.txt"
    Shell "echo ***************************************************************************************************** >> Image_Info.txt"
    Shell "echo * Below is the list of Windows editions and associated indicies available for the above named file. * >> Image_Info.txt"
    Shell "echo ***************************************************************************************************** >> Image_Info.txt"
    Shell "echo. >> Image_Info.txt"

    If GetWimInfo_Silent = 0 Then

        ShowInfo:
        Cls
        Print "Do you also want to see information for the boot.wim and winre.wim images?"
        Print "Note that to show this information it will take a while to mount the images."
        Print
        Input "Show info for the boot.wim and winre.wim? ", Temp$
        YesOrNo Temp$
        Select Case YN$
            Case "Y"
                ShowExtendedInfo$ = "Y"

                ' Ask for a location where we can mount the install.wim

                GetDestinationPath11:

                Do
                    Cls
                    Print "In order to show the information for the winre.wim, we need to mount the install.wim image first. Please specify a"
                    Print "location that we can use for this. If the folder that you specify does not already exist, then we will will try to"
                    Print "create it."
                    Print
                    Line Input "Enter the path where the project should be created: ", DestinationFolder$

                Loop While DestinationFolder$ = ""

                CleanPath DestinationFolder$
                DestinationFolder$ = Temp$ + "\"

                ' We don't want user to specify the root of a drive

                If Len(DestinationFolder$) = 3 Then
                    Cls
                    Color 14, 4
                    Print "Please do not specify the root directory of a drive."
                    Color 15
                    Print #5, ":: It appears that the root directory of a drive was specified. This is not a valid location."
                    Print #5, ""
                    Pause
                    GoTo GetDestinationPath11
                End If

                ' Check to see if the destination specified is on a removable disk

                Cls
                Print "Performing a check to see if the destination you specified is a removable disk."
                Print
                Print "Please standby..."
                DriveLetter$ = Left$(DestinationFolder$, 2)
                RemovableDiskCheck DriveLetter$
                DestinationIsRemovable = IsRemovable

                Select Case DestinationIsRemovable
                    Case 2
                        Cls
                        Color 14, 4: Print "This is not a valid disk.";: Color 15: Print " Please specify another location."
                        Print #5, ":: An invalid disk was specified."
                        Print #5, ""
                        Pause
                        GoTo GetDestinationPath11
                    Case 1
                        Cls
                        Print "The disk that you specified is a removable disk. ";: Color 14, 4: Print "Please specify a fixed disk.": Color 15
                        Print
                        Print "NOTE: Project must be created on a fixed disk due to limitations of some Microsoft utilities."
                        Print #5, ":: The specified disk is a removable disk. This is not valid."
                        Print #5, ""
                        Pause
                        GoTo GetDestinationPath11
                    Case 0
                        ' if the returned value was a 0, no action is necessary. The program will continue normally.
                End Select

                ' Verify that the path specified exists.

                If Not (_DirExists(DestinationFolder$)) Then

                    ' The destination path does not exist. We will now attempt to create it.

                    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + Chr$(34)
                    Shell _Hide Cmd$

                    ' Checking for existance of folder again again to see if we were able to create it.

                    If Not (_DirExists(DestinationFolder$)) Then
                        Cls
                        Color 14, 4: Print "The destination does not exist and we were not able to create the destination folder.": Color 15
                        Print
                        Print "Please recheck the path you have specified and try again."
                        Print #5, ":: The destination does not exist and could not be created."
                        Print #5, ""
                        Pause
                        GoTo GetDestinationPath11
                    End If
                End If

                ' If we have arrived here it means that the destination path already exists
                ' or we were able to create it successfully.

                ' Start by setting an AV exclusion for the destination path. We will log this location to a temporary file
                ' so that if the file is interrupted unexpectedly, we can remove the exclusion the next time the program
                ' is started.

                ' IMPORTANT: The count of files listed immediately below is the number of files of each type in the folders specified
                ' INCLUDING FILES THAT WILL NOT BE UPDATED.

                ' Add an AV exclusion for the destination folder

                CleanPath DestinationFolder$
                ff = FreeFile
                Open "WIM_Exclude_Path.txt" For Output As #ff
                Print #ff, Temp$
                Close #ff
                Cmd$ = "powershell.exe -command Add-MpPreference -ExclusionPath " + "'" + Chr$(34) + Temp$ + Chr$(34) + "'"
                Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

            Case "X"
                GoTo ShowInfo
        End Select

        Cls
        Print "**************************"
        Print "* Mounting the ISO image *"
        Print "**************************"
        Print
    End If

    MountISO SourcePath$
    LocalTemp$ = MountedImageDriveLetter$ + "\x64"
    Architecture$ = "SINGLE"

    ' The lines below test to see if this image has an install.esd or an install.wim and runs the appropriate command.
    ' Normally, we should not need this. Only an install.wim should be present for this project, but this routine can handle either.

    InstallFileTest$ = MountedImageDriveLetter$ + "\sources\install.wim"
    InstallFileTest2$ = MountedImageDriveLetter$ + "\sources\install.esd"
    InstallFile$ = "" ' Set an initial value

    If _FileExists(InstallFileTest$) Then
        InstallFile$ = "\sources\install.wim"
    ElseIf _FileExists(InstallFileTest2$) Then
        InstallFile$ = "\sources\install.esd"
    Else
        Cls
        Print "We were not able to find either an install.wim or an install.esd file in the specified image."
        Print "Please specify a valid image."
        Pause
        GoTo GetWimInfo_Main_Done
    End If

    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Get-WimInfo /WimFile:" + Chr$(34) + MountedImageDriveLetter$ + InstallFile$ + Chr$(34) + " /index:1 > WIM_Info2.txt"
    Shell Chr$(34) + Cmd$ + Chr$(34)

    ' Determine the Windows Build Number

    A$ = _ReadFile$("WIM_Info2.txt")

    If _FileExists("WIM_Info2.txt") Then
        Kill "WIM_Info2.txt"
    End If

    X = InStr(A$, "ServicePack Build")
    Y = _InStrRev(X, A$, ":") + 2
    Version$ = Mid$(A$, Y, ((X - Y) - 2))

    X = InStr(A$, "ServicePack Level")
    Y = _InStrRev(X, A$, ":") + 2
    SP_Build$ = Mid$(A$, Y, ((X - Y) - 2))

    Shell "echo *********************************************************************** >> Image_Info.txt"
    Shell "echo * The Windows editions in this image have the following build number: * >> Image_Info.txt"

    Cmd$ = "echo * " + Version$ + "." + SP_Build$ + Space$(68 - ((Len(Version$) + Len(SP_Build$) + 1))) + "* >> Image_Info.txt"
    Shell Cmd$
    Shell "echo *                                                                     * >> Image_Info.txt"
    Shell "echo * Note: It is assumed that all editions have the same build number    * >> Image_Info.txt"
    Shell "echo *********************************************************************** >> Image_Info.txt"

    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Get-WimInfo /WimFile:" + Chr$(34) + MountedImageDriveLetter$ + InstallFile$ + Chr$(34) + " >> Image_Info.txt"
    Shell Chr$(34) + Cmd$ + Chr$(34)

    ' If the user wanted to show info for the winre.wim and the boot.wim, gather and display that info now.
    ' If this routine is being run silently, then we do not gather extended info.

    If GetWimInfo_Silent = 1 Then GoTo SkipExtendedInfo
    If ShowExtendedInfo$ <> "Y" Then GoTo SkipExtendedInfo

    ' Get info for the boot.wim file and save it to Image_Info3.txt

    Print "*********************************"
    Print "* Getting info for the boot.wim *"
    Print "*********************************"
    Print

    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Get-WimInfo /WimFile:" + Chr$(34) + MountedImageDriveLetter$ + "\sources\boot.wim" + Chr$(34) + " /index:1 > WIM_Info3.txt"
    Shell Chr$(34) + Cmd$ + Chr$(34)

    ' Determine the boot.wim Build Number

    A$ = _ReadFile$("WIM_Info3.txt")

    If _FileExists("WIM_Info3.txt") Then
        Kill "WIM_Info3.txt"
    End If

    X = InStr(A$, "ServicePack Build")
    Y = _InStrRev(X, A$, ":") + 2
    Version$ = Mid$(A$, Y, ((X - Y) - 2))

    X = InStr(A$, "ServicePack Level")
    Y = _InStrRev(X, A$, ":") + 2
    SP_Build$ = Mid$(A$, Y, ((X - Y) - 2))

    Shell "echo. >> Image_Info.txt"
    Shell "echo ****************************************************** >> Image_Info.txt"
    Shell "echo * The boot.wim image has the following build number: * >> Image_Info.txt"

    Cmd$ = "echo * " + Version$ + "." + SP_Build$ + Space$(51 - ((Len(Version$) + Len(SP_Build$) + 1))) + "* >> Image_Info.txt"
    Shell Cmd$
    Shell "echo ****************************************************** >> Image_Info.txt"

    If InstallFile$ = "\sources\install.esd" Then
        Shell "echo. >> Image_Info.txt"
        Shell "echo Because this Windows image uses an install.esd file rather than an install.wim file, we will not show >> Image_Info.txt"
        Shell "echo winre.wim information. >> Image_Info.txt"
        Shell "echo. >> Image_Info.txt"
        GoTo SkipExtendedInfo
    End If

    Print "**********************************"
    Print "* Getting info for the winre.wim *"
    Print "* This will take a while!        *"
    Print "**********************************"
    Print

    ' Create the folders needed to mount the install.wim and copy the install.wim from the source.
    ' Start by removing any old folders by the names of "image" or "mount" to clear up any leftover junk.

    Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "image" + Chr$(34) + " /s /q"
    Shell _Hide Cmd$
    Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "mount" + Chr$(34) + " /s /q"
    Shell _Hide Cmd$

    ' Now create the new folders.

    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "image" + Chr$(34) + " >nul 2>&1"
    Shell Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "mount" + Chr$(34) + " >nul 2>&1"
    Shell Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "robocopy " + MountedImageDriveLetter$ + "\sources " + DestinationFolder$ + "image install.wim > NUL"
    Shell Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "attrib -h -s -r " + Chr$(34) + DestinationFolder$ + "image\install.wim" + Chr$(34) + " > NUL"
    Shell Chr$(34) + Cmd$ + Chr$(34)

    ' Mount the install.wim and get info for the winre.wim.

    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /mount-image /imagefile:" + Chr$(34) + DestinationFolder$ + "image\install.wim" + Chr$(34) + " /index:1 /mountdir:" + Chr$(34) + DestinationFolder$ + "mount" + Chr$(34) + " > NUL"
    Shell Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Get-WimInfo /WimFile:" + Chr$(34) + DestinationFolder$ + "mount\Windows\System32\Recovery\winre.wim" + Chr$(34) + " /index:1 > WIM_Info4.txt"
    Shell Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /unmount-image /mountdir:" + Chr$(34) + DestinationFolder$ + "mount" + Chr$(34) + " /discard" + " > NUL"
    Shell Chr$(34) + Cmd$ + Chr$(34)

    ' Remove the AV exclusion for the destination folder

    CleanPath DestinationFolder$
    Cmd$ = "powershell.exe -command Remove-MpPreference -ExclusionPath " + "'" + Chr$(34) + Temp$ + Chr$(34) + "'"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    If _FileExists("WIM_Exclude_Path.txt") Then Kill "WIM_Exclude_Path.txt"

    ' Determine the winre.wim Build Number

    A$ = _ReadFile$("WIM_Info4.txt")

    If _FileExists("WIM_Info4.txt") Then
        Kill "WIM_Info4.txt"
    End If

    X = InStr(A$, "ServicePack Build")
    Y = _InStrRev(X, A$, ":") + 2
    Version$ = Mid$(A$, Y, ((X - Y) - 2))

    X = InStr(A$, "ServicePack Level")
    Y = _InStrRev(X, A$, ":") + 2
    SP_Build$ = Mid$(A$, Y, ((X - Y) - 2))

    Shell "echo. >> Image_Info.txt"
    Shell "echo ******************************************************* >> Image_Info.txt"
    Shell "echo * The winre.wim image has the following build number: * >> Image_Info.txt"

    Cmd$ = "echo * " + Version$ + "." + SP_Build$ + Space$(52 - ((Len(Version$) + Len(SP_Build$) + 1))) + "* >> Image_Info.txt"
    Shell Cmd$
    Shell "echo ******************************************************* >> Image_Info.txt"
    Shell "echo. >> Image_Info.txt"

    ' Perform a cleanup of the folders that were needed to get the winre.wim info

    Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "image" + Chr$(34) + " /s /q"
    Shell _Hide Cmd$
    Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "mount" + Chr$(34) + " /s /q"
    Shell _Hide Cmd$

    ' This is the end of the section to extract info for the boot.wim and winre.wim images

    SkipExtendedInfo:

    If GetWimInfo_Silent = 0 Then
        Print "*************************"
        Print "* Dismounting the image *"
        Print "*************************"
        Print
    End If

    Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourcePath$ + "'" + Chr$(34) + Chr$(34) + " > NUL"
    Shell Cmd$

    GetWimInfo_Main_Done:

End Sub


Sub ProcessRangeOfNums (Range$, CheckForValidOrder)

    ' This process will break down a string into a list of numbers. Pass to it a range of index numbers in a string
    ' and a "1" if it want it to check for valid number ordering from low to high or a "0" if you don't want
    ' this check. It will return 3 values:
    '
    ' 1) Each individual number in the array RangeArray(). For example, if user specified 1-3 5 7 the array would contain 1 2 3 5 7.
    ' 2) A count of the number of elements in the integer value TotalNumsInArray. In the above example this value would be 5.
    ' 3) The Integer ValidRange would return a 0 if we detect an invalid string and a 1 if it looks like a good range of numbers.
    '
    ' Be sure to DIM all the above as SHARED. Before calling this routine, verify that you are not passing an empty string.

    ' Initialize local variables

    Dim RangeElements As Integer ' The number of elements in the TempArray() variable
    Dim x As Integer ' General counter used in FOR...NEXT loop
    Dim y As Integer ' General counter used in FOR...NEXT loop
    Dim NumbersArray(0) As String ' Holds all the numbers pulled from Range$ as strings
    Dim Seperators As String
    Dim TempArray(0) As Integer ' A temporary array holding all numbers from the Range$ passed to this routine

    ' Initialize shared variables

    ValidRange = 0
    ReDim RangeArray(0) As Integer

    ' Remove any leading or trailing spaces

    Range$ = LTrim$(Range$)
    Range$ = RTrim$(Range$)

    ' Verify that the first and last characters of the string are a number

    If ((Val(Left$(Range$, 1))) > 0) Then
        If ((Val(Right$(Range$, 1))) > 0) Or (Right$(Range$, 1) = "0") Then
            ValidRange = 1
        End If
    End If

    If ValidRange = 0 Then GoTo DoneProcessingElements

    ' Parse Range$. Each number in the string will be stored in the array called NumbersArray$() as a string. The string
    ' Seperators$ will contain all the seperators (a space or a dash in the order that they appeared in Range$. We know
    ' that the string will begin with a number, so there will always be at least one number. Once we hit a seperator (a space
    ' or a dash) there will have to be another number after that, so we will increase RangeElements to keep a count of how
    ' many numbers there are. Note that we are dimensioning NumbersArray$() to be one greater than the number of seperators
    ' because there will be one final number after the last seperator. In other words the number of numbers will be equal to
    ' (seperators + 1).

    Seperators$ = ""
    RangeElements = 1

    For x = 1 To Len(Range$)
        If Mid$(Range$, x, 1) = "-" Then
            Seperators$ = Seperators$ + "-"
            RangeElements = RangeElements + 1
        ElseIf Mid$(Range$, x, 1) = " " Then
            Seperators$ = Seperators$ + " "
            RangeElements = RangeElements + 1
        Else
            ReDim _Preserve NumbersArray(RangeElements + 1) As String
            NumbersArray$(RangeElements) = NumbersArray$(RangeElements) + Mid$(Range$, x, 1)
        End If
    Next x

    ' Verify that numbers appear in order. The second number should be greater than the first number, the third number should
    '  be greater than the second, etc.

    ' If there are no seperators, then all we have is a single number. As a result, we can skip the check for proper
    ' sequencing of the numbers.

    If Seperators$ = "" Then
        ReDim _Preserve NumbersArray(1) As String
        GoTo SequenceCheck_End
    End If

    If CheckForValidOrder = 0 Then GoTo SequenceCheck_End

    ' Start the sequencing check

    For x = 2 To RangeElements
        If Val(NumbersArray(x)) <= Val(NumbersArray(x - 1)) Then
            ValidRange = 0
            GoTo DoneProcessingElements
        End If
    Next x

    SequenceCheck_End:

    ' Expand any range of numbers into individual numbers. Convert them into integers from the string values

    ReDim TempArray(RangeElements) As Integer

    For x = 1 To RangeElements
        TempArray(x) = Val(NumbersArray$(x))
    Next x

    TotalNumsInArray = 0
    ReDim _Preserve RangeArray(1) As Integer
    RangeArray(1) = TempArray(1)
    TotalNumsInArray = 1

    For x = 1 To (RangeElements - 1)
        If Mid$(Seperators$, x, 1) = "-" Then
            For y = (TempArray(x) + 1) To TempArray(x + 1)
                TotalNumsInArray = TotalNumsInArray + 1
                ReDim _Preserve RangeArray(TotalNumsInArray) As Integer
                RangeArray(TotalNumsInArray) = y
            Next y
        Else
            TotalNumsInArray = TotalNumsInArray + 1
            ReDim _Preserve RangeArray(TotalNumsInArray) As Integer
            RangeArray(TotalNumsInArray) = TempArray(x + 1)
        End If
    Next x

    DoneProcessingElements:

End Sub


Sub AddUpdatesStatusDisplay (CurrentImage, TotalImages, StatusIndicator)

    ' This routine prints current status information for the routine to inject updates into Windows images
    ' Pass the current image number, the total number of images, and a flag to it. If the first 2 are zeros, then
    ' we print a couple of blank lines where the header would be displayed and skip the display of the header.
    ' The flag (StatusIndicator) tells us which status screen to display below (see the SELECT CASE block).

    ' Note that the status displays below (Case 1 through however many status steps there are) are not all in the
    ' order in which they occur. This is because some steps have been added and changed over time so rather than
    ' renumber all steps, some steps were added to the end, or otherwise modified.

    Dim OverallStatus As String
    Dim x As Integer

    ' Reset the display window to 120 x 30 in case this was changed so that the status disply is properly formatted.

    Width 120, 30
    Cls

    If ((CurrentImage = 0) And (TotalImages = 0)) Then
        Print
        Print
        GoTo SkipHeader
    End If

    ' Header display

    OverallStatus$ = "* Currently working on edition" + Str$(CurrentImage) + " of" + Str$(TotalImages) + " *"

    For x = 1 To Len(OverallStatus$)
        Print "*";
    Next x

    Print

    ' The section of code below will display a message if the program detects pending installs in a Windows image.

    Print OverallStatus$;
    If OpsPending$ = "Y" Then
        Print "   ";: Color 4: Print "PENDING INSTALLS DETECTED!";: Color 10: Print " More info will be displayed when routine is done.": Color 15
    Else
        Print
    End If
    For x = 1 To Len(OverallStatus$)
        Print "*";
    Next x

    ' Display program verion and progress info on the console title bar

    _ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " - Currently working on edition" + Str$(CurrentImage) + " of" + Str$(TotalImages)

    SkipHeader:

    Print
    Print

    Select Case StatusIndicator
        Case 1
            Color 0, 10: Print "- Pre-Update Task: This task is performed once before applying updates": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[             ] Mounting a Windows Edition"
            Print "[             ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[             ] Updating WinRE"
            Print "[             ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 2
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Mounting a Windows Edition": Color 15
            Print "[             ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[             ] Updating WinRE"
            Print "[             ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 3
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[             ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Updating WinRE"
            Print "[             ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 4
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[             ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only": Color 15
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 5
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[             ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only": Color 15
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Updating WinPE (Index 2 of 2)"
        Case 7
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 8
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 9
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 10
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 11
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 12
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 13
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 14
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[  COMPLETED  ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 15
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[  COMPLETED  ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[  COMPLETED  ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 16
            Color 0, 10: Print "- Pre-Update Task: This task is performed once before applying updates": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[             ] Mounting a Windows Edition"
            Print "[             ] Adding Drivers"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
        Case 17
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Mounting a Windows Edition": Color 15
            Print "[             ] Adding Drivers"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
        Case 18
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Adding Drivers"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
        Case 19
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Drivers"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
        Case 20
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Drivers"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
        Case 21
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Drivers"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
        Case 22
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Drivers"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
        Case 23
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Drivers"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[  COMPLETED  ] Moving Updated WIM File to Base Image"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Final ISO Image"
        Case 24
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Drivers"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[  COMPLETED  ] Moving Updated WIM File to Base Image"
            Print "[  COMPLETED  ] Creating Final ISO Image"
        Case 25
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Mounting a Windows Edition": Color 15
            Print "[             ] Adding Servicing Stack and Cumulative / Checkpoint / Incremental Updates"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 50
            Color 0, 10: Print "- Pre-Update Task: This task is performed once before applying updates": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[             ] Mounting a Windows Edition"
            Print "[             ] Locking in Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[             ] Updating WinRE"
            Print "[             ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 51
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Mounting a Windows Edition": Color 15
            Print "[             ] Locking in Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[             ] Updating WinRE"
            Print "[             ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 52
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[             ] Locking in Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Updating WinRE"
            Print "[             ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 53
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[             ] Locking in Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only": Color 15
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 54
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[             ] Locking in Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only": Color 15
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Updating WinPE (Index 2 of 2)"
        Case 55
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Locking in Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 56
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 57
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 58
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Base Image"
            Print "[             ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 59
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Moving Updated WIM File to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 60
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[  COMPLETED  ] Moving Updated WIM File to Base Image"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 61
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[  COMPLETED  ] Moving Updated WIM File to Base Image"
            Print "[  COMPLETED  ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied while handling the first edition being updated only"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
    End Select

    ' Display the Auto Shutdown status. If a file named "Auto_Shutdown.txt" exists on the desktop, then the system will be
    ' shutdown when the program is done running. If a file named "Auto_Hibernate" is present on the desktop then the
    ' system will hibernate. Note that this file can be created or removed / renamed by the user even while the program is
    ' running. The status will be updated in the display each time this status display routine is updated, which ocurrs
    ' with the start of each new step in the update process.
    '
    ' Also, check for the existance of a file named "WIM_PAUSE.txt" on the desktop. So long as that file exists, pause the
    ' execution of the program.

    ShutdownStatus = 0

    If _FileExists(Environ$("userprofile") + "\Desktop\Auto_Shutdown.txt") Then ShutdownStatus = ShutdownStatus + 1
    If _FileExists(Environ$("userprofile") + "\Desktop\Auto_Hibernate.txt") Then ShutdownStatus = ShutdownStatus + 2

    Select Case ShutdownStatus
        Case 0, 3
            Locate 1, 95: Print "Auto Shutdown:  ";
            Color 10: Print "Disabled";: Color 15
            Locate 2, 95: Print "Auto Hibernate: ";
            Color 10: Print "Disabled";: Color 15
        Case 1
            Locate 1, 95: Print "Auto Shutdown:  ";
            Color 14, 4: Print "Enabled ";: Color 15
            Locate 2, 95: Print "Auto Hibernate: ";
            Color 10: Print "Disabled";: Color 15
        Case 2
            Locate 1, 95: Print "Auto Shutdown:  ";
            Color 10: Print "Disabled";: Color 15
            Locate 2, 95: Print "Auto Hibernate: ";
            Color 14, 4: Print "Enabled ";: Color 15
    End Select

    Locate 29, 3: Color 0, 10: Print "  Place a file named AUTO_SHUTDOWN.TXT on desktop to shutdown system or AUTO_HIBERNATE.TXT to hibernate when done.  ";
    Locate 30, 3: Print "  Use WIM_PAUSE.TXT to pause the program. Changes are reflected when progress advances to the next step.            ";: Color 15

    ' If the file "WIM_PAUSE.txt" exists on the desktop, pause program execution until
    ' that file is deleted or renamed.

    Do While _FileExists(Environ$("userprofile") + "\Desktop\WIM_PAUSE.txt")
        Locate 3, 95: Color 14, 4: Print "PROGRAM EXECUTION PAUSED";: Color 15
        _Delay .5
        Locate 3, 95: Print "PROGRAM EXECUTION PAUSED";
        _Delay .5
    Loop

    Locate 3, 95: Print "                        ";

End Sub


Sub Pause

    ' Displays one blank line and then the message "Press any key to contine".
    ' Previously we used a simple "shell pause" in this routine, but the problem with that method is that the
    ' keyboard buffer from QB64 would not be read causing a series of pasted responses to halt when the "pause"
    ' command was run via "shell". The routine below works around this issue.

    Dim x As Integer ' Will be set to something other than "0" when keyboard input is available.
    Dim keystroke As Integer

    Print
    Print "Press any key to continue...";

    ' If a script is being recorded, add an "Enter" to the script to continue past the pause.

    If ScriptingChoice$ = "R" Then
        Print #5, ":: Pressing <ENTER> to continue after a pause."
        Print #5, "<ENTER>"
        Print #5, ""
    End If

    ' If a script is being played back, exit this subroutine since there is no need to pause when
    ' a script is being run.

    If ScriptingChoice$ = "P" Then Exit Sub

    keystroke = 0

    ' The below loop runs repeatedly until keyboard input is available

    Do
        _Limit 50 ' Linit the number of times loop is run per second to avoid putting high load on CPU

        x = 0 ' Clear x
        x = _ConsoleInput ' x will become equal to "1" if keyboard input is available

        ' The loop below is run when keyboard input is available

        If x = 1 Then
            keystroke = _CInp

            ' When a key is pressed, "keystroke" will have a value greater than "0"

            If keystroke > 0 Then

                ' If we arrive here, a key was pressed. Now we wait until the key is released

                Do

                    _Limit 50 ' Linit the number of times loop is run per second to avoid putting high load on CPU

                    x = 0
                    x = _ConsoleInput

                    If x = 1 Then
                        keystroke = _CInp
                    End If

                Loop Until keystroke < 0 ' A released key returns a negative number. When "keystroke" is a negative number, the key was released.

            End If

        End If

    Loop While keystroke = 0
    Print
End Sub


Sub DisplayFile (FileName$)

    ' Call this subroutine and pass the name of a text file to it. The file contents will be displayed to the screen
    ' one screen at a time.

    Dim FileNum As Long
    Dim x As Integer
    Dim Text As String

    FileNum = FreeFile
    Open FileName$ For Input As #FileNum

    Do Until EOF(FileNum)
        Cls
        For x = 1 To (_Height - 3)
            Line Input #FileNum, Text$
            If EOF(FileNum) Then Exit For
            Print Text$
        Next x

        ' When recording and playing back a script, the number of screens of information displayed can vary which might throw off a script. To
        ' handle this, we use the following logic:

        ' During this loop, we display the index information one screen at a time. When playing back a script, we don't want pauses after each screen of
        ' information is displayed. During the recording of a script, we want to pause so the user can see the information, but we do do not want to record
        ' the pause between screens of displayed information because this will throw off playback. As a result, we flip off recording, pause, turn recording
        ' back on again. If we have skippedrecording operations, then pause after each screen is displayed.

        Select Case ScriptingChoice$
            Case "P"
                ' Don't pause between screens of information when playing back a script
            Case "R"
                ' Temporarily flip script recording off so that the pauses when displaying screens of information is not recorded, then flip it back on.
                ScriptingChoice$ = "S"
                Pause
                ScriptingChoice$ = "R"
            Case Else
                ' Pause between screens of information. We want to pause for screens of information where we have chosen to skip scripting operations.
                ' In that case, ScriptingChoice$ will be set to "S". However, for operations where scripting is not used, even though we are not
                ' performing scripting, ScriptingChoice$ will not be set (it will be an empty string). The CASE ELSE statement covers us for both states.
                Pause
        End Select
    Loop

    Close #FileNum

End Sub


'Sub FindDISMLogErrors_SingleFile (Path$, LogPath$)


'    ' Pass to this routine the full path including a file name of a DISM log file and it will scan the file for errors.
'    ' Pass the location of the log files as the second parameter.
'    ' If an error is found it will note this and will pass back to the main program a value of "Y" in DISM_Error_Found$.
'    ' It will also create a log file called ERROR_SUMMARY.log
'    ' NOTE: This routine does not check for the existance of the file being passed to it. You should check the validity of
'    ' the path and file before calling this routine.

'    DISM_Error_Found$ = "N"
'    Dim ff As Long ' Hold the next open Free File number
'    Dim ff2 As Long ' Used to get a file number for the 2nd file that needs to be open at the same time
'    Dim Position1 As Double
'    Position1 = 0
'    Dim Position2 As Double
'    Dim StartOfError As Double
'    Dim LogFile As String
'    Dim ErrorMessage As String

'    ' Init variables

'    ff = FreeFile
'    Open Path$ For Binary As #ff
'    LogFile$ = Space$(LOF(ff))
'    Get #ff, 1, LogFile$
'    Close #ff

'    Do
'        Position1 = InStr(Position1 + 1, LogFile$, "Error                 ")
'        If Position1 Then
'            StartOfError = Position1 - 21
'            Position2 = InStr(StartOfError, LogFile$, (Chr$(13) + Chr$(10)))
'            ErrorMessage$ = Mid$(LogFile$, StartOfError, (Position2 - StartOfError))
'            DISM_Error_Found$ = "Y"
'            ff2 = FreeFile
'            Open (LogPath$ + "\ERROR_SUMMARY.log") For Append As #ff2
'            Print #ff2, "Warning! Error was reported in the log file named:"
'            Print #ff2, Path$
'            Print #ff2, "The error reported is:"
'            Print #ff2, ErrorMessage$
'            Print #ff2, ""
'            Close #ff2
'        End If
'    Loop Until Position1 = 0
'End Sub


Sub AutounattendHandling

    AskAboutAutounattend:

    Cls
    Print "Type HELP if you need information about the below option."
    Print

    Do
        Print "For this project, should we ";: Color 0, 10: Print "EXCLUDE";: Color 15: Print " the autounattend.xml answer file if it exists? ";
        Input "", ExcludeAutounattend$
    Loop While ExcludeAutounattend$ = ""

    If UCase$(ExcludeAutounattend$) = "HELP" Then
        Cls
        Print "This routine creates a bootable disk. If the file being used to create the disk includes an autounattend.xml"
        Print "answer file, you should be aware of the implications. Depending upon the configuration of the bootable disk,"
        Print "booting from this disk accidentally can cause Windows setup to run with no warning, wiping out your current"
        Print "Windows installation."
        Print
        Print "To protect against this, this routine can automatically exclude the answer file if it exists."
        Pause
        GoTo AskAboutAutounattend
    End If

    YesOrNo ExcludeAutounattend$
    ExcludeAutounattend$ = YN$

    Select Case ExcludeAutounattend$
        Case "Y", "N"
            GoTo ExcludeResponseValid
        Case Else
            Cls
            Color 14, 4: Print "Invalid response!";: Color 15: Print " Please provide a valid response."
            Pause
            GoTo AskAboutAutounattend
    End Select

    ExcludeResponseValid:

End Sub


Sub GetNumberOfIndices

    ' Determine how many x86 and x64 editions are in the image

    Dim DualArchitectureFlag As String
    Dim LastReadIndex As Integer
    Dim WimInfo As String

    ' Initialize Variables

    NumberOfSingleIndices = 0
    NumberOfx86Indices = 0
    NumberOfx64Indices = 0
    DualArchitectureFlag$ = ""
    Open "Image_Info.txt" For Input As #1
    Do
        Line Input #1, WimInfo$
        If (InStr(WimInfo$, "x86 Editions")) And (InStr(WimInfo$, "File Name:") = 0) Then DualArchitectureFlag$ = "x86_DUAL"
        If (InStr(WimInfo$, "x64 Editions")) And (InStr(WimInfo$, "File Name:") = 0) Then DualArchitectureFlag$ = "x64_DUAL"
        If Len(WimInfo$) >= 9 Then
            If (Left$(WimInfo$, 7) = "Index :") Then
                LastReadIndex = Val(Right$(WimInfo$, (Len(WimInfo$) - _InStrRev(WimInfo$, " "))))
                Select Case DualArchitectureFlag$
                    Case ""
                        NumberOfSingleIndices = LastReadIndex
                    Case "x86_DUAL"
                        NumberOfx86Indices = LastReadIndex
                    Case "x64_DUAL"
                        NumberOfx64Indices = LastReadIndex
                End Select
            End If
        End If
    Loop Until EOF(1)
    Close #1
End Sub


Sub EiCfgHandling

    ' Ask the user if they want to create an ei.cfg file

    ' IMPORTANT: This routine was originally designed to ask a user if they wanted to inject an EI.CFG file into their image. However,
    ' some users have reported that this could be confusing and have suggested that the behavior be changed to not ask about this at all
    ' and to simply create the EI.CFG file. We are now implementing that change. Note that we are keeping the original code in place here
    ' and simply bypassing it at this time. This will allow us to easily restore the original behavior should we decide to do so in the
    ' future. To restore the original functionality, do this:
    '
    ' 1) Comment out the line line below that reads "CreateEiCfg$ = "Y" by placing a single quote mark at the start of the line.
    '
    ' 2) Following that line are a series of lines that are commented out (disabled) by having a single quote mark at the start of the line.
    ' Remove the single quote marks from all the line up to just before the "End Sub" line.
    '
    ' END OF IMPORTANT NOTE


    CreateEiCfg$ = "Y"


    'EiCfg:

    'Cls
    'Print "Type ";: Color 0, 10: Print "HELP";: Color 15: Print " if you need information about the below option."
    'Print
    'Print "Do you want to inject an ";: Color 0, 10: Print "EI.CFG";: Color 15: Print " file into your final image";: Input CreateEiCfg$

    'If ScriptingChoice$ = "R" Then
    '    Print #5, ":: Do you want to inject an EI.CFG file into your final image?"
    '    If UCase$(Left$(CreateEiCfg$, 1)) = "H" Then
    '        Print #5, ":: Help for this option was requested."
    '        Print #5, "HELP"
    '    ElseIf CreateEiCfg$ = "" Then
    '        Print #5, "<ENTER>"
    '    Else
    '        Print #5, CreateEiCfg$
    '    End If
    '    Print #5, ""
    'End If

    'If UCase$(Left$(CreateEiCfg$, 1)) = "H" Then
    '    Cls
    '    Print "If you have multiple editions of Windows in an image, Windows setup may not ask you which edition to install. If your"
    '    Print "BIOS / firmware uses a signature to indicate the edition that originally shipped with the system it may simply force"
    '    Print "installation of that Windows edition. As axample, assume that you have a laptop that shipped with Windows 10 Home"
    '    Print "edition preinstalled. You upgrade the system to Windows 10 Professional. You decide that that you want to perform a"
    '    Print "clean install of Windows 10, or maybe even Windows 11. When you begin the installation you are given no choice of what"
    '    Print "Windows edition to install. Windows setup simply proceeds to install the Home edition of Windows because that is what"
    '    Print "the BIOS signature indicates was installed from the factory."
    '    Print
    '    Print "Injecting the EI.CFG file into your image will force setup to allow you to choose the edition of Windows to be"
    '    Print "installed if your image contains multiple editions."
    '    Print
    '    Print "Note that if you use an answer file to perform an unattended setup, this file will have no effect since the answer file"
    '    Print "specifies the edition to be installed."
    '    Pause
    '    GoTo EiCfg
    'End If

    'YesOrNo CreateEiCfg$
    'CreateEiCfg$ = YN$

    'If CreateEiCfg$ = "X" Then
    '    Print
    '    Color 14, 4
    '    Print "Please provide a valid response."
    '    Color 15

    '    If ScriptingChoice$ = "R" Then
    '        Print #5, ":: An invalid response was provided."
    '        Print #5, ""
    '    End If

    '    Pause
    '    GoTo EiCfg
    'End If

End Sub


Sub Skip_PE_Updates_Check

    ' Ask the user if they want to skip updates for WinPE (boot.wim)

    'IMPORTANT: A solution to the problem of updaing Windows PE has been identified and implemented. Rather than removing all the code
    ' that handles the option to skip WinPE updates, we will bypass this section. To restore  that functionality, simply comment out
    ' or remove the 2 lines noted with a comment below.

    Skip_PE_Updates$ = "N" ' Comment out or remove this line to restore the functionality of this section
    GoTo End_Skip_PE_Updates_Check ' Comment out or remove this line to restore the functionality of this section

    No_PE_Updates:

    Cls
    Print "Type HELP if you need information about the below option."
    Print
    Input "Do you want to skip updates for WinPE"; Skip_PE_Updates$

    If UCase$(Skip_PE_Updates$) = "HELP" Then
        Cls
        Print "If your updates include a Service Stack Update (SSU) and / or a Latest Cumulative Update (LCU), you can choose to"
        Print "skip these updates for Windows PE (the boot.wim file)."
        Print
        Print "IMPORTANT: Even if you choose this option, any files that you place in the PE_Files folder will still be added to"
        Print "Windows PE. This allow you to add or remove files from WinPE without having to install SSU and / or LCU updates."
        Print "Please see the documentation regarding usage of the PE_Files folder for this purpose."
        Pause
        GoTo No_PE_Updates
    End If

    YesOrNo Skip_PE_Updates$
    Skip_PE_Updates$ = YN$

    If Skip_PE_Updates$ = "X" Then
        Print
        Color 14, 4
        Print "Please provide a valid response."
        Color 15
        Pause
        GoTo No_PE_Updates
    End If

    End_Skip_PE_Updates_Check:

End Sub


Sub GetDiskDetails

    ' This subroutine will gather details about disks available in the system. After the subroutine has
    ' completed, the following information will be available:
    '
    ' NumberOfDisks - This will indicate the number of didks seen by the system
    ' DiskIDList() - This array will hold the Disk ID number for each disk. Note that this is needed because
    '     disk ID numbers are not always sequential. For example, a system may have disk ID numbers 1,2,4,5
    '     Note the missing "3".
    ' DiskDetail$() - Holds the friendly name for each disk.
    ' ListOfDisks$ - Holds output from diskpart that looks similar to this:
    '
    '  Disk ###  Status         Size     Free     Dyn  Gpt
    '  --------  -------------  -------  -------  ---  ---
    '  Disk 0    Online         7452 GB  1024 KB        *
    '  Disk 1    Online         7452 GB  1024 KB        *
    '  Disk 2    Online          931 GB  1024 KB        *
    '  Disk 3    Online         1863 GB  1024 KB        *
    '  Disk 4    Online           59 GB      0 B
    '  Disk 5    Online          238 GB   238 GB
    '
    ' IMPORTANT: The array variables are numbered sequentially, not by DiskID. Using the example above,
    '     DiskIDList(3) would be the disk with Disk ID 4 (the 3rd disk since there is no ID 3).

    Dim ff As Integer ' Used to get next free file number
    Dim ff2 As Integer ' Used to get next free file number when 2 files need to be open at the same time.
    Dim ReadLine As String ' Used to read one line at a time from file
    Dim counter As Integer ' A temporary counter variable
    Dim x As Integer ' A temporary counter variable
    Dim y As Integer ' A temporary counter variable

    ' Create a batch file run diskpart and save the output to a file

    ff = FreeFile
    Open "temp.bat" For Output As #ff
    Print #ff, "@echo off"
    Print #ff, "(echo list disk"
    Print #ff, "echo exit"
    Print #ff, ") | diskpart"
    Close #ff
    Shell "temp.bat > DiskpartOut.txt"
    Kill "temp.bat"

    ' Open the DiskpartOut.txt file and save just the portion of the file that contains the information
    ' that we are looking for.

    ff = FreeFile
    Open "DiskpartOut.txt" For Input As #ff
    ff2 = FreeFile
    Open "DiskpartOut2.txt" For Output As #ff2

    Do
        Line Input #ff, ReadLine$
    Loop Until InStr(ReadLine$, "DISKPART>")

    Do
        Line Input #ff, ReadLine$
        Print #ff2, ReadLine$
    Loop Until InStr(ReadLine$, "DISKPART>")

    Close #ff2
    Close #ff

    ' Open DiskpartOut2.txt and read it until we encounter "--------". Keep reading until a line
    ' does not have "Disk" in it. Maintain a counter to determine how many DISK Ids there are.

    ff = FreeFile
    Open "DiskpartOut2.txt" For Input As #ff

    Do
        Line Input #ff, ReadLine$
    Loop Until InStr(ReadLine$, "--------")

    counter = 0

    Do
        Line Input #ff, ReadLine$

        If Mid$(ReadLine$, 3, 4) = "Disk" Then
            counter = counter + 1
            ReDim _Preserve DiskIDList(counter)
            DiskIDList(counter) = Val(Mid$(ReadLine$, 8, 3))
        End If

        NumberOfDisks = counter
        ReDim DiskDetail$(NumberOfDisks)
    Loop Until EOF(ff)

    Close #ff
    Kill "DiskpartOut.txt"

    ' The variable "counter" now has the number of disks that diskpart reported in the system.
    ' The array "DiskIDList()" has the ID number for each disk.
    ' The total number of disks (and as a result, the array size), is equal to "NumberOfDisks".
    ' We now need to get the disk details for each Disk ID and stuff that into the array "DiskDetail()".

    For x = 1 To NumberOfDisks
        ff = FreeFile
        Open "temp.bat" For Output As #ff
        Print #ff, "@echo off"
        Print #ff, "(echo select disk"; DiskIDList(x)
        Print #ff, "echo detail disk"
        Print #ff, "echo exit"
        Print #ff, ") | diskpart"
        Close #ff
        Shell "temp.bat > DiskDetail.txt"
        Kill "temp.bat"
        ff = FreeFile
        Open "DiskDetail.txt" For Input As #ff

        For y = 1 To 11
            Line Input #ff, ReadLine$
        Next y

        DiskDetail$(x) = ReadLine$
        Close #ff
    Next x

    Kill "DiskDetail.txt"
    ListOfDisks$ = _ReadFile$("DiskpartOut2.txt")
    Kill "DiskpartOut2.txt"
    ListOfDisks$ = Left$(ListOfDisks$, InStr(ListOfDisks$, "DISKPART") - 2)
End Sub


Sub Scripting (Procedure$)

    ' Store next available file numbers in ff1 and ff2
    Dim ff1 As Integer
    Dim ff2 As Integer
    Dim LineRead As String ' Holds one line at a time read from a file

    Scripting_MakeSelection:

    ScriptingChoice$ = ""

    Do
        Cls
        Print "Do you want to ";: Color 0, 10: Print "Play";: Color 15: Print " a script, ";: Color 0, 10: Print "Record";: Color 15: Print " a script, ";:_
        Color 0, 10: Print "Skip";: Color 15: Print " scripting operations, or get ";:Color 0, 10: Print "Help";: Color 15: Print "?"
        Print
        Print "Note that you can respond with the first letter of your choice."
        Print
        Color 0, 10: Print "Play";: Color 15: Print ", ";: Color 0, 10: Print "Record";: Color 15: Print ", ";: Color 0, 10: Print "Skip";:_
        Color 15: Print ", or ";: Color 0, 10: Print "Help";: Color 15: Print ": ";
        Input "", ScriptingChoice$
    Loop While ScriptingChoice$ = ""

    ScriptingChoice$ = UCase$(Left$(ScriptingChoice$, 1))

    Select Case ScriptingChoice$
        Case "H"
            Cls
            Print "If you choose the option to record a script then your responses to the program will be saved so that the same options"
            Print "can be played back again in the future. This will save all the manual entries that you would need to make otherwise."
            Print "The script will be created with comments noting the purpose of each entry. This will allow you to easily make manual"
            Print "modifications to the script without having to go through the entire procedure of recreating the script."
            Print
            Print "The script will be saved in the same folder from which the program was run and will be called WIM_SCRIPT.TXT. You can"
            Print "rename the file if you wish to make it something more meaningful to you."
            Print
            Color 0, 10: Print "Use caution with scripting!";: Color 15: Print " If you change the contents of folders you are working with, the script may not work!"
            Pause
            GoTo Scripting_MakeSelection
        Case "P"
            Cls
            Print "What is the name of the script file that you want to run?"
            Print "Enter the full path including file name and extension."
            Print
            Line Input "Script File Name: ", ScriptFile$
            CleanPath ScriptFile$
            ScriptFile$ = Temp$

            If Not _FileExists(ScriptFile$) Then
                Cls
                Print "No such file exists in the folder where this program resides. Please specify a valid file name."
                Pause
                GoTo Scripting_MakeSelection
            End If

            ' We need to create a temporary script file that stips out all the comments and blank lines.

            ff1 = FreeFile
            Open (ScriptFile$) For Input As #ff1
            ff2 = FreeFile
            Open ("WIM_SCRIPT.TXT") For Output As #ff2

            ' Take different actions based upon what is contained in the line being read

            Do
                Line Input #ff1, LineRead$
                If Left$(LineRead$, 2) = "::" Then
                    ' This line is a comment. Ignore it
                    _Continue
                End If
                If LineRead$ = "" Then
                    ' This is a blank line. Ignore it
                    _Continue
                End If
                If LineRead$ = "<ENTER>" Then
                    ' This line is significant. We need to save an ENTER to the file.
                    Print #ff2, ""
                    _Continue
                End If
                ' This line is significant. It contains a response that needs to be saved.
                Print #ff2, LineRead$
            Loop Until EOF(ff1)

            Close #ff2
            Close #ff1
        Case "R"
            Open "WIM_SCRIPT.TXT" For Output As #5
            Print #5, "::::::::::::::::::::::::::::::::::::::::::::"
            Print #5, ":: Script to "; Procedure$; "  ::"
            Print #5, "::::::::::::::::::::::::::::::::::::::::::::"
            Print #5, ""
            Print #5, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
            Print #5, ":: A response shown as <ENTER> simply indicates that <ENTER> was pressed in response to the query.  ::"
            Print #5, ":: For a pause where user is asked to press any key to continue, we always script <ENTER> no matter ::"
            Print #5, ":: what the actual response was. When pressing <ENTER> would accept a default value, we script that ::"
            Print #5, ":: value rather than <ENTER>. When manually modifying a script, start comments with 2 colons (::).  ::"
            Print #5, ":: You can use blank lines to make the script easier to read. Blank lines are ignored.              ::"
            Print #5, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
            Print #5, ""

            ' Perform a check first to see if a file named "WIM_SCRIPT.TXT" already exists in the program folder.

            If _FileExists(ProgramStartDir$ + "\" + "WIM_SCRIPT.TXT") Then
                Cls
                Print "A file by the name of WIM_SCRIPT.TXT already exists in the folder where this program is located. We will overwrite"
                Print "this file with the script that we are about to record. If you wish to keep this file, please move it or rename it"
                Print "before you continue."
                Pause
            Else
                Cls
                Print "When the script is done recording we will save it to the same folder where this program is located with the name"
                Print "WIM_SCRIPT.TXT. You can rename it to anything that you like if you wish. You can keep your scripts anywhere, they"
                Print "do not have to be kept in the folder with the porogram."
                Pause
            End If
        Case "S"
            ' Do nothing
        Case Else
            Cls
            Print "Invalid selection"
            Pause
            GoTo Scripting_MakeSelection
    End Select
End Sub


Sub CleanVols

    ' This routine will look for any removable drives that have a status of "unusable" and removes the drive letters from them.
    ' Call this routine before performing a search for available drive letters. This is especially useful for removable media
    ' drives that have had a "clean" performed on them in diskpart. These volumes cause problems for the detection of the
    ' drive letters that they still hold.


    Dim DriveLetter As String
    Dim ff As Integer
    Dim ff2 As Integer
    Dim Temp As String
    Dim VolNum As String


    ' Send a list of volumes to a file.

    ff = FreeFile
    Open "QueryVols.bat" For Output As #ff
    Print #1, "@echo off"
    Print #1, "(echo list vol"
    Print #1, "echo exit"
    Print #1, "echo ) | diskpart > VolStatus.txt"
    Close #ff
    Shell "QueryVols.bat"

    If _FileExists("QueryVols.bat") Then Kill "QueryVols.bat"

    ff = FreeFile
    Open "VolStatus.txt" For Input As #ff

    ' Search the list of volumes for any volues that are shown as "Removable" and "Unusable". Grab the volume
    ' number and drive letter for any such volumes and use that to remove the drive letter from those volumes.

    Do
        Line Input #ff, Temp$
        If Len(Temp$) = 79 Then
            If Mid$(Temp$, 40, 9) = "Removable" Then
                If Mid$(Temp$, 61, 8) = "Unusable" Then
                    DriveLetter$ = Mid$(Temp$, 16, 1)
                    VolNum$ = RTrim$(Mid$(Temp$, 10, 2))

                    ff2 = FreeFile
                    Open "CleanVol.bat" For Output As #ff2
                    Print #ff2, "@echo off"
                    Print #ff2, "(echo select volume "; VolNum$
                    Print #ff2, "echo remove letter="; DriveLetter$
                    Print #ff2, "echo exit"
                    Print #ff2, "echo ) | diskpart > NUL"
                    Close #ff2
                    Shell "CleanVol.bat"

                    If _FileExists("CleanVol.bat") Then Kill "CleanVol.bat"

                End If
            End If
        End If

    Loop Until EOF(ff)

    Close #ff

    If _FileExists("VolStatus.txt") Then Kill "VolStatus.txt"

End Sub




' Release Notes
'
' 1.0.0.001 - July 20, 2021
' This is the first BETA release of the program. This is the original WIM Tools program but restructured to work with only Windows 11
' and not Windows 10. Technically, this program will work with Windows 10 as well but only x64 editions. In fact, this is the whole
' difference. Since Windows 11 is x64 only (64-bit) and there is no longer an x86 version (32-bit), we have removed all the portions
' of the program that handle x86 and dual architecture images. This greatly simplifies the program. Currently, this has resulted in a
' size reduction of about 30% in the code size. While I have not yet performed any performance testing, it is possible that some
' operations may be faster.
'
' This version is feature complete, but in depth testing to look for any bugs or issues still needs to be performed.
'
' 1.0.0.002 - July 20, 2021
' Since projects created by this version of the program will only have a single WIM file because no dual architecture images are used,
' we changed all references to "WIM files" to "WIM file" (singular) in putput messages.
' For the routine that creates bootable media from an ISO image, we added a check to make sure that the ISO image is x64 architecture.
' This was not needed in the original program because x86 and dual architecture images were permissable.
'
' This version of the program has also undergone testing and is the first production release of the program.
'
' 1.0.0.003 - July 21, 2021
' Very small change - Modified the icon to differentiate it from the original program.
'
' 1.0.1.004 - August 2, 2021
' Added the ability to pause program execution. For the routines that inject updates, drivers, or boot-critical drivers,
' execution can use quite a few resources, especially disk resources. When viewing the status on the screen, as progress
' goes to the next item in the list, we will check the desktop for a file named "WIM_PAUSE.txt". If such a file exists,
' execution will be paused until that file is deleted or renamed. If execution is paused a flashing message will be
' displayed to make the user aware of this.
'
' 1.0.2.005 - August 12, 2021
' Modified routines that check for elevated status. The first routine to be altered is the routine right at the start of the
' program. It was found that if the program was saved with certain characters in the file name, this routine would fail. For
' example, if program were to be renamed to "John's Program.exe", then it would fail. This corrects that situation. As an
' added benefit, this change makes the startup look better. Previously, if the program was not started elevated, the
' process of closing the original program and then relaunching it elevated caused a visible second window to be opened before
' the first one was closed. With this change that is no longer an issue.
'
' In addition, the routine that exports drivers from the system creates a batch file that also has the code to self-elevate.
' That code was also updated to resolve the same problem of not working when the batch file has certain characters in the name.
'
' 1.1.0.006 - August 16, 2021
' For the routine that will create bootable Windows installation media, we have made some improvements to make the user
' experience better. First, we create an empty text file with a unique name to the first two partitions when a "Wipe"
' operation is performed. In the future, if the user wants to perform a "Refresh" operation, we will search for these files.
' The purpose of this is to detect if more than one bootable drive created with this program exists. If more than one such
' drive is connected, then we need to ask the user for the drive letters to be refreshed. Otherwise, if only a single such
' drive is connected, we automatically detect the drive letters to be updated so that we don't need to bother the user for
' those details.
'
' 1.1.1.007 - August 16, 2021
' Added a line to the main menu to indicate the edition (Dual Architure or x64 Only). The name of the source file and the
' resulting executable have also been changed to clearly differentiate between the editions.
'
' Undid a recent change to the self-elevating code. It turns out that this was causing the code to actually NOT self-elevate.
'
' 1.1.2.008 - August 17, 2021
' After further investigation into the self-elevating code issue, we now have a more robust solution. The comments in that
' section of code explain how it works.
'
' 1.1.3.009 - August 24, 2021
' For the routine that creates bootable media, we have increased the amount of time that we wait for BitLocker to initialize
' a disk if the user chooses to create BitLocker encrypted partitions. This should ensure that even very slow media, such as
' some thumb drives, have enough time to initialize. In addition, added a block of text to better guide the user when they
' choose to create additional partitions.
'
' 2.0.0.010 - August 26, 2021
' This is a major new release. Comprehensive help has been added to the program.
'
' 2.0.1.011
' Added comments to the new help system in order to make the code easier to read and locate specific
' help topics within the code.
'
' 2.0.2.012
' Performed some tidying-up of the text in the new help section.
'
' 3.0.0.013 - September 2, 2021
' Another new major version. For the routines that inject Windows updates, drivers, or boot-critical drivers, the program can
' automatically generate a script file complete with comments. The comments make it easy to manually modify or "tweak" the
' script file. These script files can then be played back automatically by the program.
'
' 3.1.0.014 - September 2, 2021
' Added the ability still create a script without actually having to carry out an injection of updates, drivers, or boot-critical
' updates. Also improved the look of the screen asking if user wants to perform any scripting operations.
'
' 3.1.1.015 - September 2, 2021
' Added a hint to the status screens to remind user that they can use AUTO_SHUTDOWN.TXT and WIM_PAUSE.TXT files on desktop
' to perform an automatic shutdown or pause the program execution.
'
' 3.1.2.016 - September 3, 2021
' Bug fix - the final ISO image created by the routines to inject Windows updates, drivers, and boot-critical drivers was
' lacking the ".ISO" extension on the filename. This has been resolved. Also fixed a bug where script file may fail to be
' moved to the program directory after it is created.
'
' 3.1.3.017 - September 3, 2021
' Bug fix - In the help section, the organization of drivers for the routine that injects drivers did not accurately
' show the folder structure that should be used. Made a few other changes to the help for clarification purposes.
'
' 3.1.4.018 - September 7, 2021
' Confirmed for sure that there are times where the ei.cfg file is still needed. Fortunately, this was never removed from the code
' as originally planned. We simply gave a user the option to insert an ei.cfg file or not. Removed any comments from the code that
' stated that this file may no longer be needed, but leaving the functionality alone.
'
' 3.1.5.019 - September 8, 2021
' Refined the help message regarding the injection of an EI.CFG file to make the purpose of that file clearer. In addition, in other
' places in the program, when a user wants to see help, any response staring with the letter "H" would be acceptable to get help.
' This did not work for the prompt for an EI.CFG file. This has been changed to be more consistent with the rest of the program.
'
' 3.1.6.020 - September 16, 2021
' Rewrote the "pause" routine. This was a very simple routine that simply printed a blank line and then paused execution of the
' program until the user hit a key. The pause was accomplished by running the command line utility "pause" from a QB64 "shell"
' statement. For most purposes this is fine, except that it would not take input from the QB64 keyboard buffer. As a result,
' taking a series of commands and pasting them into the program would not work when a pause is encountered. This also has the
' potential to cause problems if we want to expand scripting capabilities in the future. The rewrite of this routine solves this.
'
' 3.1.7.021 - September 17, 2021
' We have 2 two "DO...LOOP" structures in the rewritten "PAUSE" code. These loops run forever waiting for a key press and release.
' Revised these structures so that they only look for a press or release 50 times per second to avoid hammering the CPU.
'
' 3.1.8.022 - September 24, 2021
' Fixed a bug that only happens when program goes back to start and reinitializes. For example, after a routine is run to completion
' and we return to the main menu. Right near the start of the program we stored the current working directory in a variable so that
' we knew the location from where the program was started. However, we then change the working directory to a temporary location
' where we can create temporary files. When jumping back to the start, we clear all variables and loose track of the original
' working directory. As a result, it is essential that we change the current working directory back to the original location before
' we jump back to the beginning of the program.
'
' Also added a note in the help for scripting to make users aware that changing contents of directories can cause script failures.
'
' Finally, the shutdown option was not working when the system was locked. The program was missing a "/f" option to force the
' shutdown. Without this, a shutdown cannot be performaed when the system is locked. We also provide a chance to abort dhutdown.
'
' 3.2.0.023 - September 26, 2021
' Made some significant changes specifically to better accomodate scripting. If the number of files in a source folder changed,
' this would break scripting because it changes your responses when the program asks if you want to update each individual file.
' To eliminate this problem, the program will now allow you to specify a full path including a file name. This way the program
' does not need to inquire about each file because you are specifying each file unambiguously.
'
' In addition, the script files are now a little easier to read and manually modify.
'
' 3.2.1.024 - September 27, 2021
' Further refinements to the scripting. The script files are much easier to read and edit. Blank lines are now legal and not interpreted
' as an <ENTER>. An <ENTER> is now explicitly shown with an "<ENTER>" tag in the script.
'
' 3.2.2.025 - September 27, 2021
' Bug fix. In the process of adjusting scripting settings, we caused a display problem for those routines that want to show index details
' but do not honor scripting settings (playback, record, or skip). This has been resolved.
'
' 3.2.3.026 - October 2, 2021
' Bug fix. At the concluion of a process, the program tries to dele the antivirus exclusion that it setup earlier and then delete a
' tracking file. It's possible that the tracking file may not exist and result in a program crash. This has been resolved.
'
' 4.0.0.027 - October 5, 2021
' This is a major new version!
' The option to create bootable media now has 2 options. The previously existing option to create a single Windows boot option along
' with multiple generic partitions to hold other data remains as was. However, we now have the option to create media that allows for
' multiple bootable partitions. For example, you could boot Windows 10 setup / recovery media, Windows 11 setup / recovery media, Macrium
' reflect recovery media, and any other Windows PE/RE based bootable media, as well as several generic partitions for holding any other
' desired data. This functionality is only for x64 / UEFI based systems and not BIOS or x86 based systems.
'
' 4.0.1.028 - October 5, 2021
' Bug fix. Even if a user did not want to hide operating system and Windows PE / RE partition drive letters, we were hiding them
' anyway. This has been fixed.
'
' 4.0.2.029 - October 11, 2021
' Minor change. Reworded some text in the routine that reorders Windows editions within an image for clarity.
'
' 4.0.3.30 - October 12, 2021
' Fixed a bug where we were skipping over an entire section of code becase we had a call to a subroutine commented out.
'
' 4.0.4.31 - October 12, 2021
' Bug fix: Still dealing with the code to create bootable media. Under a specific set of circumstances we were running the
' code to ask the user what disk write the image to twice. This has been resolved.
'
' 4.0.5.32 - October 13, 2021
' Minor change - Updated some text displayed to user for clarity.
'
' 4.1.0.33 - October 20, 2021
' Microsoft now distributes the the SSU and LCU in one combined package. We have changed our code to work with this new model.
' Note that this may look a bit odd at time becaue we apply the LCU and then apply it again. This is because the first time the
' SSU get applied, the second time, the LCU gets applied. Note that if we were loading certain optional components such as language
' packs, etc., these componts would get installed right between thos two occurences of the LCU update for both WinRE and WinPE (boot.wim).
'
' 4.1.1.34 - October 21, 2021
' Added a couple of lines to the section of code that cleans up project files. In rare circumstances, a user could abort the program
' at a point in time where cleanup of files becomes difficult. We've added a couple lines to try to even more aggressively try to
' cleanup such files.
'
' 4.1.2.35 - October 21, 2021
' Changed some wording of on screen progress indicator. With the new combined SSU and LCU updates, it appears that if an SSU exists,
' the first pass (where the SSU would be applied) completes rather quickly, and then the LCU is added on the second pass which can
' take quite a while. However, if no SSU exists, then it seems that the first pass already applies the LCU and the second pass
' completes almost instantly since there is nothing left to be done. The wording of the progress screens is now changed to reflect
' the fact that the SSU and LCU are combined and we simply call the 2 passes "pass 1" and "pass 2".
'
' 4.1.3.36 - October 23, 2021
' Changed the wording of a few messages in the routine to extract the contents of .CAB files for clarity.
'
' 4.1.4.37 - October 25, 2021
' Changed wording on a menu item.
'
' 4.1.5.38 - November 24, 2021
' For the option to display WIM image information, we have added a few lines at the end of the output to show the build number of the
' Windows editions in the image. Help has been updated to note that this information is displayed.
'
' 4.1.6.39 - January 11, 2022
' For the routine that injects updates, after the user enters a filename or path for the source, we need to determine if what was
' entered is a filename or a path. There was a bug in this logic causing a fault in the program if a path was entered. This is now
' resolved.
'
' 4.1.7.40 - January 12, 2022
' When updates are being injected into Windows images, the screen that shows the current progress is very much reliant upon being
' properly sized as 120 x 30. If the user has either accidentally or purposely changed the size of the program window, the status
' screen will not look very good at all. As a result, each time we refresh the status display, we will now reset the screen to
' a size of 120 x 30.
'
' 4.2.0.41 - January 13, 2022
' We had previously introducednew functionality to the routine for creating bootable media to allow for the creation of media that
' could boot multiple Operating Systems and / or WinPE / WinRE based media such as various recovery disks, etc. This routine is
' being removed at this time, but may be added back at a later date. There are several reasons for this:
'
' 1) We have seen several occurences of BSODs with no clear understanding yet of the cause.
' 2) The result looks sloppy - there is no nice boot menu and the resulting boot is also inconsistent on different systems. On some
'    systems you will see a boot menu item for each partition (both FAT32 and NTFS) while some systems show only the FAT32
'    partitions. It just looks sloppy and seems a little half-baked when compared to the rest of this program.
'
' 4.2.1.42 - January 14, 2022
' Permanently removed the option that we described as having disabled in the Jan 13, 2022 release notes. It's clear that if we are
' ever to bring back this functionality a major rewrite of that code would be needed.
'
' 4.2.2.43 - January 17, 2022
' Rewording of a menu item and help text for improved clarity.
'
' 4.2.3.44 - January 24, 2022
' Added the ability to easily switch the program between applying separate SSU updates or combined LCU / SSU updates.
' At the start of routine to inject updates, if we set ProcessSeparateSSU to 1, then we process a separate SSU update.
' If set to 0, then we process a combined LCU / SSU update.
'
' 4.2.4.45 - January 25, 2022
' When injecting Windows updates, there was a logic flaw. At the end of the process, we were moving the install.wim to the final
' destination. However, this results in a significantly larger file. Rather than simply moving it, we should be exporting each
' index to the destination. We correctly perform a cleanup operation on the install.wim, however, the cleanup has no effect on
' the size until an export is performed. So, as a result of neglecting to perform the export, we miss out on the benefit of the
' cleanup that was performed earlier. This has now been corrected.
'
' 4.2.5.46 - January 25, 2022
' Found that there was still a discrepancy in the size of final ISO image created by a batch file that I was using and this
' program. Noted that there was a significant difference in the boot.wim size. That led to the discovery that a variable I
' was using was wrong in some places. There was both a SSU_Update_Avail$ and an SSU_Updates_Avail$. This has been corrected.
' In addition, we were moving the original boot.wim to the final location rather than the updated file. This too has been corrected.
'
' 4.3.0.47 - February 10, 2022
' If in an effort to finally put the issue of how to properly apply combined LCU / SSU updates, we have performed so testing. Microsoft
' documentation says that SSU updates should be applied first, followed by certain other updates, and then the LCU should be applied.
' From my testing, it looks like this information is now outdated. For an online installation of Windows this may still apply, but it
' seems that for an offline servicing of a Windows image, the following should be done:
'
' 1) For WinRE.WIM - Apply the combined LCU / SSU update at the time where the SSU would normally be applied.
' 2) For BOOT.WIM (WinPE) - Apply the combined LCU / SSU update at the time where the LCU was previously applied and apply nothing where the
'    SSU was previously applied.
' 3) For INSTALL.WIM (the main OS) - Apply the combined LCU / SSU update at the time where the LCU was previously applied and apply nothing
'    where the SSU was previously applied.
'
' 4.3.1.48 - February 12, 2022
' Noticed recently that performing a "clean" within diskpart on a flash drive will often fail the first time it is run. Oddly, it always works
' the second time. It seems that this affects not only manually running the command, but affects this program as well. As a result, we simply
' run the "clean" command twice.
'
' 4.3.2.49 - February 13, 2022
' Came accross a program that needs to access the .WIM files in a Windows image. For some silly reason, this program is sensitive to the case
' of the file names, wanting only lowercase characters. There are places in the program where we were saving these file names using uppercase
' characters. This has been changed to use all lowercase.
'
' In addition, we have made a change to all ISO image creation routines. The timestamp of all files added to the image will now be set to the
' time at which the creation of the image was started. This will allow for easy identification of when an image was created.
'
' 4.3.3.50 - February 18, 2022
' Further refinement to the "clean" operation on media that we are writing to. There is a problem where performing a "clean" in diskpart will
' sometimes fail the first time. It usually works the second time but in testing with batch files we have seen failures even on the second
' time. The failures only seem to happen on MBR disks. As a result, we are performing a clean operation twice, attempting to set it to GPT
' each time. Finally, after a third clean we set it to the final desired state of either GPT or MBR.
'
' 4.3.4.51 - March 16, 2022
' Very minor changes to wording of some text in help. No functional changes in this update.
'
' 4.3.5.52 - March 17, 2022
' Another minor change: The routine to display Windows image information is one of the few pieces in this program that was intended to be able
' to handle both images that contain an install.wim file as well as those that use an install.esd file. We were not handling this correctly
' This has now been corrected. Along with this change we changed the temporary test file that was used to store information from WIM_Info.txt
' to Image_Info.txt.
'
' 5.0.0.53 - March 18, 2022
' Major new release: Added a new feature to the menu. There are a lot of people who have ISO images using an install.esd rather than an
' install.wim. There are occasions where the release of an image with .wim may be delayed by a few days and only the .esd version of the
' image is available. As a result, we now have the ability to convert an image with an install.esd into an image with an install.wim.
' In this x64 only edition, we only support conversion of single architecture images.
'
' 5.0.1.54 - March 19, 2022
' Revised the help menu so that help topics align with the actual menu item numbers. Previously, the numbers were offset by one. For
' example, menu item 1 was help topic 2.
'
' In addition, for this x64 only edition of the program, the versioning information was not updated causing wrong version information
' to be displayed.
'
' 5.0.2.55 - March 20, 2022
' Very minor update: In the built-in help, under the topic for the routine that injects drivers into Windows images, on the third page of
' the help topic "5) Organizing update files", there is a section for "SSU". This topic has been removed since Microsoft has now switched
' to a combined LCU / SSU model. This means that an SSU folder is no longer needed and a seperate SSU does not need to be downloaded.
'
' 5.0.3.56 - March 21, 2022
' At the last moment prior to compiling the last release, we accidentally introduced a bug causing the final image to not be created in the
' new routine that converts ESD images into WIM images. This has been fixed.
'
' 5.0.4.57 - March 21, 2022
' This update contains no change in functionality.  A few messages in the help sections were updated to make the messages clearer.
'
' 5.0.5.58 - March 23, 2022
' Completely revised the logging to make it far simpler. Logs were sometimes being generated that were 600MB+ in size and this was
' simply getting out of hand. We have also made one small change to the behavior of the routine that converts ESD files to WIM. Since
' we are simply performing a conversion and not altering the image in any other way, we are no longer altering the timestamp of files
' when we create the final ISO image.
'
' 5.1.0.59 - March 26, 2022
' In a little further refinement of the logging, we cleaned up some messages related to logging. In addition, we have enhanced the auto
' shutdown capability by adding the ability to hibernate rather than a shutdown. This is implemented by the user placing either an
' auto_shutdown.txt or an auto_hibernate file on the desktop.
'
' In addition, we found that when updates were injected into images by running a script, some important messages at the end of the
' program execution could be skipped since script playback skips all pauses where messages would normally be displayed to the user.
' This has been corrected.
'
' 5.1.1.60 - March 28, 2022
' No functionality changes. Previously, the initial variable declarations separated the variables declared as SHARED from those that were
' not. For sake of ease, these were merged with the other variables. Now, to find a variable declaration we simply need to find it in the
' list alphabetically since variables were already alphabetically arranged. We no longer need to look in more than one place.
'
' 5.1.2.61 - April 15, 2022
' Made some changes to the startup of the program prior to the main menu being displayed. We now determine if the current version of the
' program has ever been run before. If not, we display a message to the user encouraging them to review online help. In addition, we
' instruct them to run the program elevated if self-elevation fails due to their system settings. Also fixed a small bug where a file
' named Win_Info.txt was specified but it should have been Image_Info.txt. This did not not result in any functional issues, simply
' an extremely minor cosmetic issue.
'
' 5.1.3.62 - May 9, 2022
' No functional changes. Marking as stable build following extensive testing.
'
' 5.1.4.63 - May 23, 2022
' No functional changes. Corrected a comment that had an incorrect date and recompiled with the latest QB64pe version (0.7.1).
'
' 6.0.0.64 - June 8, 2022
' Major new version. Added the ability to create a bootable disk that can contain as many different ISO images as you wish. These
' can include various versions of Windows as well as Windows PE and Windows RE based media.

' 6.0.1.65 - June 11, 2022
' Performed much work on the code that was just added to allow creation of multi boot images. Rather than being a whole entire
' separate section of code, this is merged with the code to create a single image boot media. This gives us all the capabilities
' that the single image code had such as being able to add additional partitions, choose BitLocker encryption, etc.
'
' 6.0.2.66 - June 12, 2022
' Made some major changes to help topics covering the sections that were updated in the past few builds.
'
' 6.0.3.67 - June 13, 2022
' Determined a way to allow an override and create GPT media for media created to boot from multiple different ISO images. The
' code now implements this putting a multi image boot media on parity with the single images and providing all the same capabilities.
'
' 6.0.4.68 - June 15, 2022
' Fixed a section of code in which responding with a path with spaces but not enclosed in quotes was not being parsed correctly.
'
' 6.0.5.69 - June 17, 2022
' The program was not working with paths and filenames that contain commas. This has been corrected.
'
' 6.0.6.70 - July 3, 2022
' When booting from a disk created as Windows multi image boot disk, under certain circumstances, it was possible that a couple of
' "Access Denied" messages could be displayed and cause the user to question if everything was okay. Everything was working fine,
' but this issue has been addressed to eliminate that concern.

' 6.0.7.71 - July 6, 2022
' Found a logic bug in testing. When creating a multi image boot disk, we create a batch file that the user can run to restore the
' original state of the media. We had that batch file auto delete itself after being run. This was faulty logic. That batch file
' should not be deleted.
'
' 6.1.0.72 - July 7, 2022
' Major rework of the routine to create a bootable disk with a single Windows image or multiple Windows images plus WinPE and WinRE
' based media. Made a lot of changes to messages, fixed a lot of spelling / sentence structure issues, cleaned up status displays,
' made interaction with the user more friendly, and more.
'
' 6.1.1.73 - July 11, 2022
' No change in functionality. Simply cleaned up a user message.
'
' 6.2.0.74 - July 21, 2022
' Performed an overhaul of the routine that creates a VHD and deploys Windows to it. This includes a bug fixe and much better
' communication to the user of steps that they need to take.
'
' 6.2.1.75 - July 23, 2022
' Corrected a single spelling error.
'
' 6.2.2.76 - August 6, 2022
' Removed the last line of the "Restore_Disk.bat" batch file. This line read "END" which is an invalid keyword.
'
' 6.2.3.77 - August 14, 2022
' Minor update: For the routine "Make or update a bootable drive from one or more Windows / WinPE / WinRE ISO images", when
' all operations have been completed, we display usage instructions to the user. With each screen of instructions, the screen
' is dismissed by pressing any key. If any keys were accidentally pressed while the operations are in progress, these will be
' retained in the keyboard buffer and will then cause the messages at the end to be dismissed. In order to make sure that the
' user will see the messages, we are now clearing the keyboard buffer before we display these messages.
'
' 6.3.0.78 - August 14, 2022
' Added a minor update to flush the keyboard buffer for several other sections of code, just as was done for the
' "Make or update a bootable drive from one or more Windows / WinPE / WinRE ISO images" routine.
'
' 6.3.0.79 - August 25, 2022
' No code changes. The only change made is that we are compiling on the newly released QB64PE 3.0.
'
' 6.3.1.80 - September 2, 2022
' Minor change: When exporting drivers for a system, we created a batch file that could be used later to install
' the drivers. At the end of that batch file, we display a message to the user and pause. We no longer do this.
' This allows us to call that batch file from another batch file without it pausing and holding up execution.
'
' 6.3.2.81 - September 6, 2022
' In the routine to create a universal, bootable UFD, we made a change to create a batch file that gets saved on volume 2.
' That batch file allows to UFD to be custom configured without having to boot from the UFD. We also updated the text that
' is displayed after the UFD is created to explain how this batch file can be used.
'
' 6.3.3.82 - September 11, 2022
' Fixed some flaws in the routine to inject boot critical drivers.
'
' 6.3.4.83 - September 30, 2022
' Rearchitected parts of the routine to create a bootable disk that allows for the selection from multiple ISO images. The
' routine now creates a single batch file that can be used to check the status of the disk, to select an image to make bootable,
' or to revert the disk back to the state where no image is yet selected.
'
' 6.3.5.84 - October 1, 2022
' For robocopy commands, we often have error checking code to make sure that we have not run out of space on the destination. However,
' the error correction is never triggered because robocopy tries over and over after a 30 second delay. To resolve this, we have added
' a /r:0 to the robocopy commands to tell robocopy not to retry. This will allow the code to fall through to the error checking code.
'
' 6.3.6.85 - October 2, 2022
' Minor tweak: In the routine that creates a universal bootable disk, we had to copy the "choice" command to the disk because it is
' not present in Windows PE. We have modified the code so that this command is no longer needed.
'
' 6.4.0.86 - October 20, 2022
' The method used to inject Servicing Stack Updates (SSU) has been completely rearchitected. Microsoft listed a known issue in the
' latest Windows patches for which a workaround is listed. That issue affects only people who inject updates into Windows images.
' The workaround describes how to extract the SSU from the combined SSU / LCU update package. Previously, Microsoft did not describe
' how to do this. As a result, we were using a method devised by trial and error. This method works, but we are now applying the SSU
' using the method described by Microsoft.
'
' 6.4.1.87 - October 31, 2022
' Cleared the keyboard buffer at the end of a few more code sections to make sure that messages to the user are not overrun.
'
' 6.4.2.88 - November 9, 2022
' In the routine to create a universal boot disk that will allow for selection from multiple ISO images, we have eliminated the
' need for a user to provide a temporary location to build the WinPE image. In addition, the message that is displayed to a user
' the first time the program is run, suggesting the review of help and instructing what to do if the program cannot self-elevate
' has been revised. It will no longer be displayed after new versions of the program are installed. It will now truly be shown
' only one time.
'
' 6.4.3.89 - December 7, 2022
' Made some changes to the help file contents to make to information presented easier to understand.
'
' 6.4.4.90 - January 9, 2023
' No code changes other than these comments. Simply moving to the new QB64PE version 3.5.0 to compile the code.
'
' 6.5.0.91 - January 10, 2023
' While working with the mitigation for a Microsoft vulnerability related to BitLocker and the Windows Recovery Environment, it was
' discovered that Microsoft documentation regarding how WinRE.wim should be updated may not be correct. Although documented otherwise,
' it appears that applying the LCU to the WinRE.wim may in fact be necessary even though all indications from Microsoft are that the
' LCU does not apply to WinRE.wim. The program has been updated to apply the LCU in addition to the updates that are already being
' applied (the SSU and the SafeOS Dynamic Update).
'
' 6.5.1.92 - February 8, 2023
' For the routine that displays basic WIM information, we have added the ability to display the build number of both the boot.wim
' and winre.wim images.
'
' 6.6.0.93 - February 9, 2023
' Updated the online help to reflect the changes made in version 6.5.1.92.
'
' 6.6.1.94 - March 7, 2023
' Remove the message that is shown the first time running the program. This message unexpectedly gets displayed again whenever the temp
' folder gets cleaned out, and the message is simply proving more annoying than useful.
'
' 7.0.0.95 - March 31, 2023
' Add the ability to update the registry settings in the BOOT.WIM to bypass Win 11 system requirement checks.
'
' 7.1.0.96 - April 5, 2023
' When making physical boot media, we default to creating the first partition with a size of 2.5 GB. Often, this is way more space than is
' needed. However, it is enough space that it can cause the second partition to be too small to hold even a single Windows image on an 8 GB
' disk. To resolve this, we now give the user an option to specify the size of the first partition.
'
' 7.1.1.97 - April 12, 2023
' For the routine that injects registry changes into the BOOT.WIM to bypass Windows 11 system requirements, if a user opted to copy the
' final file over the source file, replacing the original source after the update, we would then delete the entire project file. We
' have modified this behavior so that any ISO files that were originally located in that location would be left in place. This allows
' the user to use the same project folder for multiple operations such as injecting Windows updates, without having to remember to move
' the resulting ISO file every time for fear of the next operation deleting it.
'
' 7.1.2.98 - May 1, 2023
' Minor change to some text for better readability.
'
' 7.1.3.99 - May 3, 2023
' For the routine that creates a multiboot disk from which the user can select the image to be booted, we have made a few changes. If the
' selected image contains an autounattend.xml answer file, we no longer move it to the second volume. We leave this file on the first volume.
' By doing this, the answer file is deleted when we revert the disk back to the original state, rendering the disk safe to boot from. Note
' that it is possible that a user could manually place an autounattend.xml answer file on the second volume. For safety, when we revert the
' volume back to the original state, we delete this answer file if it is present and display a note to the user that any answer file that
' may have been present has been deleted.
'
' 7.1.4.100 - May 8, 2023
' Made a number orf refinements and enhancements to the routine that creates a VHD and deploys a new instance of Windows to it.
'
' 7.1.5.101 - May 8, 2023
' Made a number of refinements to how we find available drive letters. In addition, we should now be able to correctly handle corner cases
' where detecting drive letters in use by removable media that was BitLocker encrypted but has had a "clean" operation performed on it, did
' not always work properly.
'
' 7.2.0.102 - June 16, 2023
' Made many refinements to the routine that creates Windows bootable media and multi image boot media.
'
' 7.2.1.103 - June 18, 2023
' More refinements to the routine that creates Windows bootable media and multi image boot media.
'
' 7.2.2.104 - August 5, 2023
' Completely rewrote the help that describes how to organize Windows updates for projects where updates are injected into Windows image(s).
'
' 7.2.2.105 - August 5, 2023
' In the new routine to inject Windows system requirement bypasses into the BOOT.WIM, it was realized that the first prompt for the image
' to be updated was not clear. This has been updated to specifically ask for the Windows ISO image file in order to avoid confusion.
'
' 7.2.2.106 - August 7, 2023
' Very minor change. Corrected text describing the purpose of the ei.cfg file.
'
' 7.2.2.107 - August 8, 2023
' Minor change. Changed text that read "to updated" to "to be updated".
'
' 7.2.2.108 - August 10, 2023
' Made some minor changes when asking a user if they want to include an EI.CFG file as it was reported that this option may be confusing.
'
' 7.2.2.109 - August 14, 2023
' Minor change: Clarified the prompt to the user for the location of the Windows update files to clarify that the user does not need
' to include the x64 or x86 folders in the path.
'
' 7.2.2.110 - September 14, 2023
' Very minor change: There is no change in functionality. We simply changed the wording of the main menu item for creating a VHDX to indicate
' that Win 11 23H2+ includes this capability in the GUI. We also updated the help for this feature.
'
' 7.2.3.111 - September 22, 2023
' For the routine that creates a multiboot disk that allows the user to select from a list of ISO images to make bootable, we have found
' that if a previously created boot.wim file was used, we would sometimes fail if the system was booted from the media without a boot
' selection having been made prior to the boot. This is because search the media for the volume label but the volume label that we
' embedded in the boot.wim may not be valid for the current project. This has been corrected so that we no longer look for the volume
' label. We now look for the tags named "VOL1_M_MEDIA.WIM" and "VOL2_M_MEDIA.WIM" instead.
'
' 7.3.0.112 - September 25, 2023
' Modified the code that allows creation of a multi image boot disk to display a list of available unattended answer files and allow the user
' to select an answer file to use if they wish to do so. Eliminated the need to check volume labels in the bootable media in order to be able
' to perform a refresh operation.
'
' 7.3.0.113 - November 2, 2023
' Corrected an issue that was causing the boot.wim updates to not be properly applied.
'
' 7.3.0.114 - November 21, 2023
' Revised the procedure for syncing mismatched files. It appears that only setup.exe files now need to be synced.
'
' 7.3.0.115 - December 2, 2023
' Updated the help to reflect recent changes in the program. In addition, when the users views the status of a multi boot disk by running
' the Config_UFD.bat file, we now report what answer file is configured for use if one was selected.
'
' 7.3.0.116 - December 3, 2023
' Made a number of refinements to the batch files created ny the multi disk boot reoutine.
'
' 7.3.0.117 - December 5, 2023
' Made a very minor change to the command line for OSCDIMG when creating a generic ISO image.
'
' 7.3.0.118 - December 6, 2023
' Made some further enhancements to the batch file used to configure a multi ISO image boot disk.
'
' 7.3.0.119 - January 3, 2024
' Updated the help section to reference a new type of update - the OOBE ZDP update.
'
' 7.4.0.120 - January 18, 2024
' Microsoft released some information that has led to the most recent changes. First, it was discovered that even though Microsoft
' now includes SSUs in the same package with the LCU, there can be rare circumstances where a Standalone SSU may be released. The
' program is now prepared to handle this. In addition, after injecting all the Windows updates, there is a step where we sync
' files that may be out of sync between Windows PE and the media outside of any WIM. Updated the sync procedure to include all
' needed files.
'
' 7.4.0.121 - January 21, 2024
' Previously, based upon a WinRE vulnerability from January 2023, we had added a step to apply the LCU to WinRE. However, this step
' is now no longer needed so we are removing that step from the update process.
'
' 7.4.1.122 - March 5, 2024
' Very minor change: There are several spots in the program where we read the entire contents of a file into a string. This process
' requires about 5 lines of code. QB64 now has a new "_READFILE$" command that performs all 5 lines of code in a single file. More
' importantly, it is so much easier to read. As a result, we have replaced all occurrences of the old code with this new command.
' This makes no change to the functionality of the program, just cleans things up in a few spots in the program.
'
' 7.5.0.123 - April 11, 2024
' Made a few changes to the routine that deploys Windows to a VHD for native boot. No change in functionality has been made, only some
' messages were changed for the sake of clarity. We have also added a message asking the user to NOT deploy the VHD to a drive that is
' encrypted with BitLocker.
'
' 7.5.0.124 - April 18, 2024
' Further refinement to messaging for the tool that installs a dual-boot Native VHD installation of Windows.
'
' 7.6.0.125 - May 8, 2024
' With the latest mitigations for the BlackLotus UEFI Bootkit released in April of 2024, it is necessary to apply an additional patch
' after the disk has been created. This patch needs to be applied from a system that already has these mitigations applied. As a result,
' we will check to see if the system on which we are running has these mitigations applied. If the mitigations are applied, we will
' automatically apply the patch. If the mitigations are not applied, we will suggest to the user that this patch be applied from a
' system that has the mitigations installed.
'
' PENDING - The code implemented today has not been test with dual artitecure images. It's possible that some of the commands we are
' using are not placing files in the correct location on dual architecture media. This needs to be looked into. We are placing this
' as a low priority since we are moving away from all dual architecture support. If it works as is, that's just a lucky bonus.
'
' In the section that creates a generic ISO image file, we were specifying the use of UDF version 2.0. This seems to sometimes cause
' problems. As a result, we stepped this back to UDF version 1.02.

' 7.7.0.126 - May 24, 2024
' Ironically, we are removing the changes that were just made. This because the new ADK and Windows PE versions for Windows 24H2 have
' now been released and these no longer require patching.
'
' 7.8.0.127 - June 7, 2024
' With the latest Windows Insider Release Preview ISO image (build 26100.560), physical media created by this program would not run
' correctly. We have discovered that by creating a copy of the "boot", "efi", and "support" folders on the second partition of the
' bootable media, we resolve this problem. These folders are very small (currently about 42MB), so this should pose no operational
' difficulties.
'
' 7.8.0.128 - July 18, 2024
' Minor change: Removed the registry entries for BypassCpuCheck. This no longer appears to be needed. Note that we have commented out
' those lines rather than removing them entirely just in case the need for these entries ever reappears.
'
' 7.8.1.129 - July 23, 2024
' Found that when changing the NAME and DESCRIPTION metadata for a Windows edition, updating these fields for more than one image does
' not work. This has been resolved. The problem was that we were missing a double quote mark (") at the end of the command to change the
' name and description metadata.
'
' 7.8.1.130 - August 12, 2024
' No functional changes to the program were made. This update is merely to note that we have switched to the new version of the QB64PE
' compiler (version 3.14).
'
' 23.0.0.261 - August 26, 2024
' No functional changes to the program, however, we come to the point in the program where we finally transition completely to x64 only
' support. As a result, we are adopting a new versioning scheme for the x64 only version to bring it into alignment with the scheme used
' for the dual architecture version of the program.
'
' 24.0.0.262 - October 8, 2024
' Finally! We have working 24H2 support. It turns out that one additional file needs to be copied from the boot.wim to the install.wim
' file that was not needed previously. That file is "setuphost.exe".
'
' 24.0.0.263 - October 10, 2024
' Since Windows 11 24H2 introduces Checkpoint and Incremental updates in place of the previous cumulative updates, the status screens
' needed to be updated. We now reflect that we are applying servicing stack updates and cumulative / checkpoint / incremental updates.
' We have also updated HELP to reference checkpoint / incremental updates.
'
' 24.1.0.264 - October 13, 2024
' Added a new capability: We can now set the tag for the "Installation Type" of the Windows editions in an image to "Server". This
' eliminates the need to run "setupprep /product server" when performing an upgrade installation on unsupported hardware. By setting
' this tag you can perform an upgrade installation on unsupported hardware without needing to do anything; no commands need to be run,
' no registry entries need to be set, nothing at all needs to be done. It will just work. For this same routine, we have updated the
' progress display to be clearer. Finally, the HELP system has been updated to reflect the recent changes.
'
' 24.1.1.265 - October 25, 2024
' For the routine that adds bypasses for systems that do no meet Win 11 requirements, the initial ISO image name is always Windows.ISO.
' We have changed this to ask the user for the name to use.
'
' 24.1.1.266 - December 15, 2024
' No code changes, created a new version simply to recompile under the new QB64PE 4.0.0 compiler.
'
' 24.2.0.267 - January 19, 2025
' Updated the routine that converts an install.esd into an install.wim so that it can also convert an install.wim into an install.esd.
'
' 24.3.0.268 - January 20, 2025
' Further updated the routine that allows conversion between Install.WIM and Install.ESD files. Rather than only being able to convert
' these files only from an ISO image, we can now point to an individual file as well.
'
' 24.3.1.269 - February 24, 2025
' Add clarification to messages indicating that a valid x64 folder does not exist to clarify that why an x64 subfolder is needed.
'
' 24.3.2.270 - February 26, 2025
' Windows 24H2 has introduced a new setup experience. Under investigation now is a problem where injecting Windows updates will
' sometimes cause a clean installation to fail. For a manual installation this is easy to overcome because the setup GUI provides an
' option to run the previous setup. However, this is a problem when run unattended setup. We have added the ability to the routine that
' patches Windows so that it can be installed on technically unsupported hardware to also optionally allow forcing the use of the previous
' setup. This will cause both manual and unattended setup to use the previous setup, but it is really only necessary for unattended setup.
'
' 24.3.2.272 - Feruary 27, 2025
' Very minor update. This update makes no functional changes, just changes the messaging around the updates made yesterday. When we added
' the functionality to force Windows to use the previous version of setup, we believed that there might be an issue with the program. That
' turned out not to be the case. Days after the problematic Windows update, additional supporting updates (Safe OS and Setup Dynamic
' Updates were released. So, technically, nothing in the program needed to be changed. On the plus side we have gained the functionality of
' being able to force the previous version of setup to be used.
'
' 24.3.3.273 - March 5, 2025
' Added the ability to add boot critical drivers to the boot.wim used to create bootable Windows media. Note that this only affects media that
' is created to allow the user to select from multiple images to boot from. This is because the the standard type of boot media is created from
' an ISO image file, so if additional drivers are needed, these should be injected into that ISO image first.
'
' 24.3.4.274 - April 4, 2025
' This release has no change in functionality. We merely corrected some text in a message that was not properly worded.
'
' 24.3.4.275 - April 29, 2025
' Minor change: Corrected a path that is incorrectly displayed. The path is displayed with two backslashes (\\) where only one should be shown.
' Note that this has ne effect on functionality of the program. This is just a minor cosmetic issue.
'
' 24.3.5.276 - June 19, 2025
' Minor change: For the "Other" folder where we place .NET FX updates as well as any other updates that do not fit into other categories,
' we were searchiong for iles with only the .MSU extension. I UUP update was released today that had the .NET update as a .CAB file rather
' than the usual .MSU. This shows a flaw in logic here - since anything could be put in this folder, we should really be searching for all
' files rather than just files with a .MSU extension.
'
' 24.3.6.277 - June 26, 2025
' For the routines that inject updates into a Windows ISO image, keystrokes that were made following a script or made while the program
' was running might cause the final status screen to be dismissed when switching to it. This has been resolved.
'
' 24.3.7.278 - July 11, 2025
' After creating a multiboot disk that allows the user to select an ISO image to make bootable, we display several screens to explain the
' usage of the disk that was created. We have updated some text in this area for improved clarity. No functional changes were made in this build.
'
' 25.0.0.279 - August 21, 2025
' Add a major new feature - an unattended answer file generator.
'
' 25.0.1.280 - August 24, 2025
' Added a batch file version of the unattended answer file generator to the batch file that configures a multiboot disk. So now the user has
' the option to use previously made answer files or generate one on the fly.
'
' 25.0.1.281 - August 25, 2025
' Updated the help section to reflect the recent chages. Fixed a bug in the new code that could cause the Config_UFD.bat file to crash
' if no answer file was present in the "Anser Files" folder of a multi boot disk create with this program.


