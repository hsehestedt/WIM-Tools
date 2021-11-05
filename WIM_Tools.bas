' WIM (Windows Image Manager) Tools, Dual Architecture Edition
' (c) 2021 by Hannes Sehestedt

' Release notes can be found at the very end of the program

' This program is intended to be run on Windows 10 x64 or Windows 11 and should be compiled with the 64-bit version of QB64.
' IMPORTANT: It is very important to use the March 17, 2021 Dev Build or newer of QB64 as this fixes a bug with the CLEAR command.


Option _Explicit

' ********************************************************
' ** Make sure to keep the "$VersionInfo" updated below **
' ********************************************************

' Perform some initial setup for the program

Rem $DYNAMIC
$ExeIcon:'iso.ico'
$VersionInfo:CompanyName=Hannes Sehestedt
$VersionInfo:FILEVERSION#=19,1,4,182
$VersionInfo:ProductName=WIM Tools Dual Architecture Edition
$VersionInfo:LegalCopyright=(c) 2021 by Hannes Sehestedt
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

' This program needs to be run elevated. Check to see if the program is running elevated, if not,
' restart in elevated mode and terminate current non-elevated program.

' We need to begin by parsing the original command line for a "'" (single quote) character.
' If this character exists, it will cause the command to self-elevate the program to fail.
' We resolve this by changing the single quote character into two single quotes back-to-back.
' This will "escape" the single quote character in the command.

Temp1$ = Command$(0)
For x = 1 To Len(Temp1$)
    If Mid$(Temp1$, x, 1) = "'" Then
        Temp2$ = Temp2$ + "''"
    Else
        Temp2$ = Temp2$ + Mid$(Temp1$, x, 1)
    End If
Next x


If (_ShellHide(">nul 2>&1 " + Chr$(34) + "%SYSTEMROOT%\system32\cacls.exe" + Chr$(34) + " " + Chr$(34) + "%SYSTEMROOT%\system32\config\system" + Chr$(34))) <> 0 Then
    Shell "powershell.exe " + Chr$(34) + "Start-Process '" + (Mid$(Temp2$, _InStrRev(Temp2$, "\") + 1)) + "' -Verb runAs" + Chr$(34)
    System
End If

' If we reach this point then the program was run elevated.


' ***********************************************************
' ***********************************************************
' ** The following strings hold the program version number **
' **               and program release date.               **
' **            Make sure to keep this updated.            **
' ***********************************************************
' ***********************************************************

Dim Shared ProgramVersion As String ' Holds the current program version. This is displayed in the console title throughout the program
Dim ProgramReleaseDate As String

ProgramVersion$ = "19.1.4.182"
ProgramReleaseDate$ = "Oct 25, 2021"


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
Dim AnswerFilePresent As String
Dim Arc As String ' Used to store architecture type in the routine to create a VHD
Dim Architecture As Integer ' Flag that gets set to 1 for a single architecture image, 2 for dual architecture, and 0 if an invalid image
Dim ArchitectureChoice As String ' Search code for ArchitectureChoice$ for a comment explaining usage
Dim AutoSize As String ' Set to "Y" if the last partition is to be autosized to occupy all remaining space
Dim AvailableSpace As Long ' Used for tracking available space on a disk
Dim AvailableSpaceString As String ' Used for tracking available space on a disk
Dim bcd_ff As Integer ' Holds a free file number for file access to BCD files
Dim BitLockerCount As Integer ' Stores the number of partitions that need to be encrypted
Dim CDROM As String ' The drive letter assigned to the mounted ISO image
Dim ChosenIndex As Integer
Dim Column As Integer ' Used for positioning cursor on screen
Dim CurrentImage As Integer ' A counter used to keep track of the image number being processed
Dim CurrentIndex As String ' A counter used to keep track of the index number within an image being processed
Dim CurrentIndexCount As Integer
Dim Description As String ' Holds the description metadata to be assigned to a Windows edition within an image
Dim DescriptionFromFile As String ' Holds the DESCRIPTION field of an image parsed from WIM_Info.txt file
Dim DestArcFlag As String ' A flag that varies wit the architecture type used to build out a final path
Dim Destination As String ' Destination path
Dim DestinationFileName As String ' The file name of the ISO image to be created without a path
Dim DestinationFolder As String ' The destination folder where all the folders created by the project will be located as well as the final updated ISO images
Dim DestinationIsRemovable As Integer ' Flag to indicate if the originally specified destination is removable
Dim DestinationPath As String ' The destination path for the ISO image without a file name
Dim DestinationPathAndFile As String ' The full path including the file name of the ISO image to be created
Dim DiskID As Integer ' Used in multiple places to ask the user for a DiskID as presented by the Microsoft DiskPart utility
Dim DiskIDSearchString As String ' Holds a disk ID that will be searched for in the output of diskpart commands
Dim DisplayUnit As String ' Holds "MB", "GB", or "TB" to indicate what units user is entering partition size in
Dim DriveLetter As String ' Take a path and store the drive letter from that path (C:, D:, etc.) in this valiable to be used to determine if drive is removable or not
Dim DST As String ' A path that includes location of install.wim files
Dim DualArchitectureFlag As String
Dim DualArcImagePath As String
Dim DualBootPackage As Integer
Dim DUALFileCount As Integer
Dim EditionName As String
Dim ErrMsg(4) As String
Dim ErrMsgFile1 As Long
Dim ErrMsgFile2 As Long
Dim exFATorNTFSdriveletter As String
Dim ExportFolder As String ' Used by the routine for exporting drivers from a system as well as the Reorg routine.
Dim FAT32DriveLetter As String ' Letter assigned to 1st partition
Dim ff As Long ' Holds the value returned by FREEFILE to determine an available file number
Dim ff2 As Long
Dim FileLength As Single
Dim FileSourceType As String
Dim FinalImageName As String
Dim FSType As String 'Set to either NTFS of EXFAT to determine what filesystem user wants to use
Dim HideLetters As String ' Set to "Y" to indicate that drive letter for this partition should be removed
Dim Highest_Single As Integer
Dim Highest_x64 As Integer
Dim Highest_x86 As Integer
Dim IDX As String
Dim ImageInfo As String
Dim ImagePath As String
Dim ImageSourceDrive As String
Dim Index As String ' Holds index number for the image being processed as a string without leading space
Dim IndexCountLoop As Integer
Dim IndexOrder As String
Dim IndexRange As String ' A temporary string for a user to specify a range of numbers. Example: 1-3 5 7-8
Dim IndexString As String ' This is the value of the integer variable Index converted to a string
Dim IndexVal As Integer
Dim InjectionMode As String ' From the main menu, set to "UPDATES" if user wants to inject Windows updates, or "DRIVERS" if user wants to inject drivers.
Dim InstallFile As String
Dim InstallFileTest As String
Dim LCU_Updates_Avail As String
Dim LettersAssigned As Integer
Dim MainLoopCount As Integer ' Counter to indicate which loop we are in.
Dim MakeBootablePath As String
Dim MakeBootableSourceISO As String ' The full path and file name of the ISO image that the user want to make a bootable thumb drive from
Dim ManualAssignment As String
Dim MaxLabelLength As Integer ' The allowable length for a volume label - 11 for exFAT, 32 for NTFS
Dim MediaLetter As String
Dim MenuSelection As Integer ' Will hold the number of the menu option selected by the user
Dim MoreFolders As String
Dim MountDir As String ' Used to hold text while reading from a file looking for a DISM mount location
Dim Multiplier As Single
Dim NameFromFile As String ' Holds the NAME field of an image parsed from WIM_Info.txt file
Dim NewLabel As String
Dim NumberOfx64Updates As Integer
Dim NumberOfx86Updates As Integer
Dim Offset As Integer
Dim OpsPendingFileCheck As String
Dim OS_Count As Integer ' The number of operating systems to be added to an image (Note that number of partitions for OS will be 2 per OS)
Dim OS_Partitions As Integer ' The number of Windows operating system partitions to be created on a GPT boot disk
Dim Other_Partitions As Integer ' The number of non-bootable partitions to be created on a GPT boot disk
Dim Other_Updates_Avail As String
Dim OutputFileName As String ' For Windows multiboot image program, holds the final name of the ISO image to be created (file name and extension only, no path)
Dim Override As String
Dim ParSizeInMB As String ' Holds the size of a partition as a string
Dim ParSize(0) As Long ' For a GPT boot media project, holds the size of each partition being created.
Dim ParType(0) As String ' For a GPT boot media project, holds the file system type of each partition being created.
Dim Par1InstancesFound As Integer
Dim Par2InstancesFound As Integer
Dim PartitionCounter As Integer ' used as a counter when processing partitions
Dim PartitionDescription(0) As String ' Friendly description for each partition in a GPT boot media project
Dim PE_Files_Avail As String
Dim PE_Partitions As Integer ' The number of Windows PE based program partitions to be created on a GPT boot disk
Dim ProjectArchitecture As String ' In Multiboot program, hold the overall project architecture type (x86, x64, or DUAL)
Dim ProjectType As String
Dim ReadLine As String
Dim ReorgFileName As String
Dim ReorgSourcePath As String
Dim RowEnd As Integer
Dim Row As Integer ' Used for positioning cursor on screen
Dim SafeOS_DU_Avail As String
Dim Setup_DU As String ' Holds the location of the Setup Dynamic update file
Dim SourceFolder As String ' Will hold a folder name
Dim SourceFolderIsAFile As String ' If SourceFolder$ actually contains a filename rather than a path, set this to "Y", else set to "N"
Dim SourceImage As String
Dim SourcePath_Multi(0) As String ' Holds the source ISO image path in a GPT boot media project
Dim SSU_Update_Avail As String
Dim TempLong As Long
Dim TempPath As String ' A temporary variable used while manipulating strings
Dim TempValue As Long
Dim TotalImages As Integer
Dim TotalImagesToUpdate As Integer
Dim TotalIndexCount As Integer
Dim TotalPartitions As Integer ' The total number of partitions that need to be created on a bootable thumb drive
Dim TotalSpaceNeeded As Long
Dim Silent As String
Dim SingleImageTag As String
Dim SingleImageCount As Integer
Dim SourceArcFlag As String
Dim SourcePath As String ' Holds the path containing the files to be injected into an ISO image file
Dim SRC As String
Dim SSU_Updates_Avail As String
Dim TempPartitionSize As String
Dim TempUnit As String
Dim TotalFiles As Integer
Dim Units As String
Dim UpdateAll As String ' User is asked by routine to inject updates or drivers if all images should be updated. This string hold their response
Dim UpdatesLocation As String
Dim UpdateThisFile As String
Dim UserCanPickFS As String
Dim UserSelectedImageName As String ' If the user wants a specific name for the final ISO image, this string will hold that name
Dim ValidDisk As Integer ' Set to 0 if user chooses an invalid Disk ID, 1 if their choice is valid
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
Dim WimInfoFound As Integer ' A flag used to indicate whether an index specified by user was found successfully in WIM_Info.txt file
Dim WipeOrRefresh As Integer
Dim x64ExportCount As Integer
Dim x86ExportCount As Integer
Dim x64_IndexOrder As String
Dim x64FileCount As Integer
Dim x64ImageCount As Integer
Dim x64Updates As String
Dim x64UpdateImageCount As Integer ' Same as x64Images except this is the number of actual images to be updated
Dim x86UpdateImageCount As Integer ' Same as x86Images except this is the number of actual images to be updated
Dim x86_IndexOrder As String
Dim x86FileCount As Integer
Dim x86ImageCount As Integer
Dim x86Updates As String
Dim y As Integer ' General purpose loop counter
Dim z As Integer ' General purpose loop counter

' Variables dimensioned as SHARED (Globally accessible to the main program and SUB procedures)

Dim Shared CleanupSuccess As Integer ' Set to 0 if we do not successfully clean contents of folder in the cleanup routine, otherwise set to 1 if we succeed
Dim Shared CreateEiCfg As String ' Set to "Y" or "N" to indicate whether an ei.cfg file should be created.
Dim Shared DiskDetail(0) As String ' Stores details about each disk in the system
Dim Shared DiskIDList(0) As Integer ' Used to store a list of valid Disk ID numbers
Dim Shared DISM_Error_Found As String ' Holds a "Y" if an error is found in log file, a "N" if not found.
Dim Shared DISMLocation As String ' Holds the location of DISM.EXE as reported by the registry
Dim Shared ErrorsWereFound As String ' Set to "N" initially, but if any errors are found in the log files, then we set this to "Y" and inform the user after processing.
Dim Shared ExcludeAutounattend As String ' If set to "Y" then exclude any existing autounattend.xml file, if set to "N" then it is okay to copy the file
Dim Shared FileCount As Integer ' The number of ISO image files that need to be processed. In multiboot image creation program, this hold the number of images we have to process.
Dim Shared ImageArchitecture As String ' Used by the DetermineArchitecture routine to determine if an ISO image is x86, x64, or dual architecture
Dim Shared IMAGEXLocation As String ' The location of the ImageX ADK utility as reported by the registry
Dim Shared IsRemovable As Integer ' Value returned subroutine to determine if a disk ID or drive letter passed to it is removable or not
Dim Shared ListOfDisks As String
Dim Shared NumberOfFiles As Integer ' Used by the FileTypeSearch subroutine to keep count of the number of files found in a folder of the type specified by a user
Dim Shared MountedImageCDROMID As String ' The MountISO returns this value
Dim Shared MountedImageDriveLetter As String ' The MountISO returns this value
Dim Shared NumberOfDisks ' Stores the number of disk drives that diskpart sees in the system
Dim Shared NumberOfSingleIndices As Integer
Dim Shared NumberOfx64Indices As Integer
Dim Shared NumberOfx86Indices As Integer
Dim Shared OpsPending As String
Dim Shared OSCDIMGLocation As String ' Holds the location of OSCDIMG.EXE as reported by the registry
Dim Shared ProgramStartDir As String ' Holds the original starting location of the program
ReDim Shared RangeArray(0) As Integer ' Each individual numeric value from the range of numbers passed into the ProcessRangeOfNums routine expanded into individual numbers
Dim Shared ScriptingChoice As String ' Used to track what scripting operation user wishes to perform
Dim Shared ScriptContents As String ' Holds the entire contents of previously created script for playback
Dim Shared ScriptFile As String ' Used to store the name of the script file to be run
Dim Shared Skip_PE_Updates As String ' if Set to "Y" we will not apply SSU and LCU updates to the WinPE (boot.wim) image
Dim Shared Temp As String ' Temporary string value that can be shared with subroutines and also used elsewhere as temporary storage
Dim Shared TempLocation As String ' This variable will hold the location of the TEMP directory.
Dim Shared TotalNumsInArray As Integer
Dim Shared ValidRange As Integer ' A flag that indicates whether a range of numbers supplied by a user is valid or not. 0=Invalid, 1=Valid
Dim Shared YN As String ' This variable is returned by the "YesOrNo" procedure to parse user response to a yes or no prompt. See the SUB procedure for details

' Arrays

Dim Shared TempArray(100) As String ' Used by FileTypeSearch subroutine to keep the name of each file of type specified by user. We assume that we will need less than 100.

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
' WinRE_x86_Present()
' x64Array()
' x86Array()

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

' If a temporary file named "WIM_File_Copy_Error.txt" still exists from a pevious un of the pogram, delete it.
' This makes no functional difference to the program, it simply cleans up a file that is not needed.

If _FileExists("WIM_File_Copy_Error.txt") Then Kill "WIM_File_Copy_Error.txt"

' If a file by the name of "WIM_Shutdown_log.txt" exists, this means that on the last run of the program
' the user chose to perform a shutdown after the program finished. Before the shutdown the program saves
' any status messages to this file. We will now display that information to the user and then delete the
' file.

If _FileExists("WIM_Shutdown_log.txt") Then
    Cls
    Print "We have detected that the last time this program was run, it was requested that the system be shutdown after the"
    Print "program completed. Prior to shutdown, we saved any status messages that you may need to be aware of. These"
    Print "messages are displayed below."
    Color 10
    Print "_________________________________________________________________________________________________________________"
    Color 15
    Print
    Shell "Type WIM_Shutdown_log.txt"
    Print
    Color 10
    Print "_________________________________________________________________________________________________________________"
    Color 15
    Print
    Print "End of Messages"
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
Print " Dual Architecture Edition         "
Print " Version "; ProgramVersion$; "                "
Print " Released "; ProgramReleaseDate$; "             "
Color 15
Print
Print
Print
Color 0, 14
Print "    1) Inject Windows updates into one or more Windows editions and create a multi edition bootable image       "
Print "    2) Inject drivers into one or more Windows editions and create a multi edition bootable image               "
Print "    3) Inject boot-critical drivers into one or more Windows editions and create a multi edition bootable image "
Color 0, 10
Print "    4) Make or update a bootable drive from one or more Windows ISO images and Windows PE / RE images           "
Print "    5) Create a bootable Windows ISO image that can include multiple editions                                   "
Print "    6) Create a bootable ISO image from Windows files in a folder                                               "
Print "    7) Reorganize the contents of a Windows ISO image                                                           "
Color 0, 3
Print "    8) Get WIM info - display basic info for each WIM in an ISO image                                           "
Print "    9) Modify the NAME and DESCRIPTION values for entries in a WIM file                                         "
Color 0, 6
Print "   10) Export drivers from this system                                                                          "
Print "   11) Expand drivers supplied in a .CAB file                                                                   "
Print "   12) Create a Virtual Disk (VHDX)                                                                             "
Print "   13) Create a VHD, deploy Windows to it, and add it to the boot menu to make a dual boot configuration        "
Print "   14) Create a generic ISO image and inject files and folders into it                                          "
Print "   15) Cleanup files and folders                                                                                "
Color 15
Color 0, 13
Print "   16) Program help                                                                                             "
Color 0, 8
Print "   17) Exit                                                                                                     "
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
Locate 5, 40: Color 0, 13
Print "   ";
Color 15
Print " Help"
Locate 27, 0
Input "   Please make a selection by number (17 Exits from the program): ", MenuSelection

' Some routines require that the Windows ADK be installed. We will now check to see if the option selected by the user is one of those routines.
' If it is, then we warn the user and return them to the main menu. If the ADK was found, then we skip this check.

If ADKFound = 1 Then GoTo Skip_ADK_Check

Select Case MenuSelection
    Case 1, 2, 3, 5, 7, 11, 12, 13, 14, 15
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
            Open "WIM_SCRIPT.TXT" For Binary As #5
            ScriptContents$ = Space$(LOF(5))
            Get #5, 1, ScriptContents$
            Close #5
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
            Open "WIM_SCRIPT.TXT" For Binary As #5
            ScriptContents$ = Space$(LOF(5))
            Get #5, 1, ScriptContents$
            Close #5
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
            Open "WIM_SCRIPT.TXT" For Binary As #5
            ScriptContents$ = Space$(LOF(5))
            Get #5, 1, ScriptContents$
            Close #5
            Kill ("WIM_SCRIPT.TXT")
            _ScreenPrint ScriptContents$
        End If
        EiCfgHandling
        GoTo InjectUpdates
    Case 4
        AutounattendHandling
        EiCfgHandling
        GoTo MakeBootDisk
    Case 5
        ExcludeAutounattend$ = "Y"
        EiCfgHandling
        GoTo MakeMultiBootImage
    Case 6
        GoTo MakeBootDisk2
    Case 7
        GoTo ChangeOrder
    Case 8
        GoTo GetWimInfo
    Case 9
        GoTo NameAndDescription
    Case 10
        GoTo ExportDrivers
    Case 11
        GoTo ExpandDrivers
    Case 12
        GoTo CreateVHDX
    Case 13
        GoTo AddVHDtoBootMenu
    Case 14
        GoTo CreateISOImage
    Case 15
        GoTo GetFolderToClean
    Case 16
        GoTo ProgramHelp
    Case 17
        GoTo ProgramEnd
End Select

' We arrive here if the user makes an invalid selection from the main menu

Cls
Color 14, 4
Print
Print "You have made an invalid selection.";
Color 15
Print " You need to make a selection by entering a number from 1 to 17."
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
' For each file, determine if it is dual architecture, x86, or x64.

' The variable InjectionMode$ will be set to "UPDATES" if we are injecting Windows updates, and "DRIVERS" if we are injecting drivers

' Initialize variables

DISM_Error_Found$ = ""
ErrorsWereFound$ = ""
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
    Input "Enter the path: ", SourceFolder$

    If ScriptingChoice$ = "R" Then
        Print #5, ":: Path to one or more Windows images or full path with a file name:"
        If SourceFolder$ = "" Then
            Print #5, "<ENTER>"
        Else
            Print #5, SourceFolder$
        End If
        Print #5, ""
    End If

Loop While SourceFolder$ = ""

'if the path ends with .ISO then we need to determine if this path is a file name or a folder name.

If UCase$(Right$(SourceFolder$, 4)) = ".ISO" Then

    If _DirExists(SourceFolder$) Then
        ' The name specified is a legit folder name
        SourceFolderIsAFile$ = "N"
        GoTo FolderNameOK
    End If

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
        Print "Please standby for a moment. Verifying the architecture of the following image:"
        Print
        Color 10
        Print FileArray$(TotalFiles)
        Color 15
        Print
        Temp$ = FolderArray$(TotalFiles) + FileArray$(TotalFiles)
        DetermineArchitecture Temp$, 1
        Select Case ImageArchitecture$
            Case "x64", "x86"
                FileSourceType$(TotalFiles) = ImageArchitecture$
            Case "DUAL"
                FileSourceType$(TotalFiles) = "x64_DUAL"
                TotalFiles = TotalFiles + 1

                ' Init variables

                ReDim _Preserve UpdateFlag(TotalFiles) As String
                ReDim _Preserve FileArray(TotalFiles) As String
                ReDim _Preserve FolderArray(TotalFiles) As String
                ReDim _Preserve FileSourceType(TotalFiles) As String

                UpdateFlag$(TotalFiles) = "Y"
                FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
                FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
                FileSourceType$(TotalFiles) = "x86_DUAL"
            Case "NONE"
                Cls
                Color 14, 4: Print "WARNING!";: Color 15: Print " An invalid file has been selected."
                Print "Check the following file to make sure that it is valid. It needs to contain INSTALL.WIM file(s), not INSTALL.ESD."
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

            ' Init variables

            ReDim _Preserve UpdateFlag(TotalFiles) As String
            ReDim _Preserve FileArray(TotalFiles) As String
            ReDim _Preserve FolderArray(TotalFiles) As String
            ReDim _Preserve FileSourceType(TotalFiles) As String

            UpdateFlag$(TotalFiles) = "Y"
            FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
            FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
            Cls
            Print "Please standby for a moment. Verifying the architecture of the following image:"
            Print
            Color 10
            Print FileArray$(TotalFiles)
            Color 15
            Print
            Temp$ = FolderArray$(TotalFiles) + FileArray$(TotalFiles)
            DetermineArchitecture Temp$, 1
            Select Case ImageArchitecture$
                Case "x64", "x86"
                    FileSourceType$(TotalFiles) = ImageArchitecture$
                Case "DUAL"
                    FileSourceType$(TotalFiles) = "x64_DUAL"
                    TotalFiles = TotalFiles + 1

                    ' Init variables

                    ReDim _Preserve UpdateFlag(TotalFiles) As String
                    ReDim _Preserve FileArray(TotalFiles) As String
                    ReDim _Preserve FolderArray(TotalFiles) As String
                    ReDim _Preserve FileSourceType(TotalFiles) As String

                    UpdateFlag$(TotalFiles) = "Y"
                    FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
                    FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
                    FileSourceType$(TotalFiles) = "x86_DUAL"
                Case "NONE"
                    Cls
                    Color 14, 4: Print "WARNING!";: Color 15: Print " An invalid file has been selected."
                    Print "Check the following file to make sure that it is valid. It needs to contain INSTALL.WIM file(s), not INSTALL.ESD."
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

' For any dual architecture images, we need to determine if the user wants to update Windows editions
' in both the x64 and x86 folders.

For x = 1 To TotalFiles
    If (FileSourceType$(x) = "x86_DUAL") Or (FileSourceType$(x) = "x64_DUAL") Then
        Cls
        Print "The file listed below is a dual architecture image. We need to know if you intend to update both x64 and x86"
        Print "editions of Windows or not."
        Print
        Color 10: Print "Filename: "; FileArray$(x): Color 15
        Print
        Print "Do you want to update ";: Color 0, 10: Print "ANY";: Color 15: Print " of the ";: Color 0, 14: Print Left$(FileSourceType$(x), 3);
        Color 15: Print " editions within this file";: Input Temp$

        If ScriptingChoice$ = "R" Then
            Print #5, ":: Do you want to update ANY of the "; Left$(FileSourceType$(x), 3); " editions within this file:"
            Print #5, "::    Filename: "; FileArray$(x)
            If Temp$ = "" Then
                Print #5, "<ENTER>"
            Else
                Print #5, Temp$
            End If
        End If

        YesOrNo Temp$
        Select Case YN$
            Case "X"
                Cls
                Print
                Color 14, 4
                Print "Please provide a valid response."
                Color 15
                If ScriptingChoice$ = "R" Then
                    Print #5, ":: An invalid response was provided."
                    Print #5, ""
                End If
                Pause
            Case "N"
                UpdateFlag$(x) = "N"
        End Select
    End If
Next x

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
    If (FileSourceType$(IndexCountLoop) = "x64_DUAL") Or (FileSourceType$(IndexCountLoop) = "x86_DUAL") Then
        Print "*******************************************************"
        Print "* This file is a dual architecture file. Please enter *"
        Print "* the index numbers for the ";: Color 0, 14: Print ">> "; Left$(FileSourceType$(IndexCountLoop), 3); " <<";: Color 15: Print " editions below. *"
        Print "*******************************************************"
        Print
    End If
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

    ' We have to specify the values for the ArchitectureChoice$ if our source is a dual architecture file so that
    ' the path to the install.wim or install.esd is correct.

    Select Case FileSourceType$(IndexCountLoop)
        Case "x64_DUAL"
            ArchitectureChoice$ = "x64"
        Case "x86_DUAL"
            ArchitectureChoice$ = "x86"
        Case Else
            ArchitectureChoice$ = ""
    End Select

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
        Select Case ArchitectureChoice$
            Case ""
                Temp$ = _Trim$(Str$(NumberOfSingleIndices))
                IndexRange$ = "1-" + Temp$
            Case "x64"
                Temp$ = _Trim$(Str$(NumberOfx64Indices))
                IndexRange$ = "1-" + Temp$
            Case "x86"
                Temp$ = _Trim$(Str$(NumberOfx86Indices))
                IndexRange$ = "1-" + Temp$
        End Select
        If IndexRange$ = "1-1" Then IndexRange$ = "1"
    End If
    Kill "WIM_Info.txt"

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

        EndProcessRange:

    End If

    ' We will now get WIM info and save it to a file called WIM_Info.txt. We will parse that file to verify that the index
    ' selected is valid. If not, we will ask the user to choose a valid index.

    SourcePath$ = FolderArray$(IndexCountLoop) + FileArray$(IndexCountLoop)
    Print
    Print "Verifying indices."
    Print
    Print "Please standby..."
    Print
    GetWimInfo_Main SourcePath$, 1

    ' If we are processing a file from a dual architecture image, then we need to make sure that we are only processing the
    ' part of the file that pertains to the x64 or the x86 portion of the file that we need.

    For x = 1 To TotalNumsInArray
        WimInfoFound = 0 ' Init Variable
        DualArchitectureFlag$ = ""
        Open "WIM_Info.txt" For Input As #1
        Do
            Line Input #1, WimInfo$
            If (InStr(WimInfo$, "x86 Editions")) Then DualArchitectureFlag$ = "x86_DUAL"
            If (InStr(WimInfo$, "x64 Editions")) Then DualArchitectureFlag$ = "x64_DUAL"
            If (FileSourceType$(IndexCountLoop) = "x86_DUAL") Or (FileSourceType$(IndexCountLoop) = "x64_DUAL") Then
                If FileSourceType$(IndexCountLoop) <> DualArchitectureFlag$ Then GoTo SkipToNextLine_Section1
            End If
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
    Kill "WIM_Info.txt"

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
    Input "Enter the path where the project should be created: ", DestinationFolder$
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

' Get a count of x64, x86, and DUAL architecture images. If we have both x64 and x86 images but no
' dual architecture images, then we will build the base image from a combination of files taken from both
' the x64 and x86 images and we dynamically build the "bcd" files that are unique to a dual architecture
' image. If we find a dual architecture image, then save the path for this image to the variable
' DualArcImagePath$ so that we can use it for building the base image.

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

x86FileCount = 0
x64FileCount = 0
DUALFileCount = 0
DualArcImagePath$ = ""

' The next set of variables will hold the actual number of each image type to be processed

x86UpdateImageCount = 0
x64UpdateImageCount = 0

For x = 1 To TotalFiles
    Select Case FileSourceType$(x)
        Case "x64"
            x64FileCount = x64FileCount + 1
            If UpdateFlag$(x) = "Y" Then x64UpdateImageCount = x64UpdateImageCount + IndexCount(x)
        Case "x86"
            x86FileCount = x86FileCount + 1
            If UpdateFlag$(x) = "Y" Then x86UpdateImageCount = x86UpdateImageCount + IndexCount(x)
        Case "x86_DUAL", "x64_DUAL"
            DUALFileCount = DUALFileCount + 1
            If (UpdateFlag$(x) = "Y") And (FileSourceType(x) = "x64_DUAL") Then x64UpdateImageCount = x64UpdateImageCount + IndexCount(x)
            If (UpdateFlag$(x) = "Y") And (FileSourceType(x) = "x86_DUAL") Then x86UpdateImageCount = x86UpdateImageCount + IndexCount(x)
            If DUALFileCount = 1 Then DualArcImagePath$ = FolderArray$(x) + FileArray$(x)
    End Select
Next x

' NOTE: When updating the image count, dual architecture images will count as 2 images since we
' list the x64 and x86 images seperately. As a result, when done getting the count, we will need to
' divide the count for dual architecture images by 2.

DUALFileCount = DUALFileCount / 2
TotalImagesToUpdate = x64UpdateImageCount + x86UpdateImageCount

' Create a flag to indicate if this project will be a single architecture project or dual architecture.

If ((x64UpdateImageCount > 0) And (x86UpdateImageCount > 0)) Then
    ProjectType$ = "DUAL"
Else
    ProjectType$ = "SINGLE"
End If

If ProjectType$ = "SINGLE" Then GoTo END_GetDualArcImagePath

' If DualArcImagePath is blank, this means that we do not have a dual architecture image available
' and we will create the base image using files from the x64 and x86 images and dynamically create
' the "bcd" files.

If DualArcImagePath$ <> "" Then
    DualBootPackage = 0
    GoTo END_GetDualArcImagePath
Else
    DualBootPackage = 1
    GoTo END_GetDualArcImagePath
End If

END_GetDualArcImagePath:

' Get the location for x64 and x86 updates or drivers to be injected

GetUpdatesLocation:

x64Updates$ = "" 'Set initial value

' If x64UpdateImageCount = 0 Then GoTo End_GetUpdatesLocation

If InjectionMode$ = "UPDATES" Then
    Do
        Cls
        Input "Enter the path to the Windows update files: ", UpdatesLocation$
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
        Input "Enter the path to the drivers: ", UpdatesLocation$
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
        Input "Enter the path to the boot-critical drivers: ", UpdatesLocation$
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
x86Updates$ = UpdatesLocation$ + "\x86"

If x64UpdateImageCount = 0 Then GoTo End_Getx64UpdatesLocation

' Verify that the x64 path specified exists.

If Not (_DirExists(x64Updates$)) Then

    ' The path does not exist. Inform user and allow them to try again.

    Cls
    Color 14, 4: Print "The specified x64 folder does not exist.";: Color 15: Print " Please try again."
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

        FileTypeSearch (x64Updates$ + "\LCU\"), ".MSU", "N"
        NumberOfx64Updates = NumberOfx64Updates + NumberOfFiles

        FileTypeSearch (x64Updates$ + "\Other\"), ".MSU", "N"
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

' Process the x86 updates location information

' Verify that the path specified exists.

If x86UpdateImageCount = 0 Then GoTo End_Getx86UpdatesLocation

If Not (_DirExists(x86Updates$)) Then

    ' The path does not exist. Inform user and allow them to try again.

    Cls
    Color 14, 4: Print "The specified x86 folder does not exist.";: Color 15: Print " Please try again."
    If ScriptingChoice$ = "R" Then
        Print #5, ":: The specified x86 folder does not exist."
        Print #5, ""
    End If
    Pause
    GoTo GetUpdatesLocation
End If

' If we have arrived here it means that the path is valid.

' Now, verify that update files actually exist in this location.
' In the case where we are applying Windows updates, we're just
' going to make sure that there is a .MSU file(s)in the LCU subfolder
' of the location specified. For driver updates, we're going to make
' sure that .INF file(s) exist.

Select Case InjectionMode$
    Case "UPDATES"

        NumberOfx86Updates = 0 ' Set initial value

        FileTypeSearch (x86Updates$ + "\LCU\"), ".MSU", "N"
        NumberOfx86Updates = NumberOfx86Updates + NumberOfFiles

        FileTypeSearch (x86Updates$ + "\Other\"), ".MSU", "N"
        NumberOfx86Updates = NumberOfx86Updates + NumberOfFiles

        FileTypeSearch (x86Updates$ + "\Setup_DU\"), ".CAB", "N"
        NumberOfx86Updates = NumberOfx86Updates + NumberOfFiles

        FileTypeSearch (x86Updates$ + "\SafeOS_DU\"), ".CAB", "N"
        NumberOfx86Updates = NumberOfx86Updates + NumberOfFiles

        FileTypeSearch (x86Updates$ + "\PE_Files\"), "*", "N"
        NumberOfx86Updates = NumberOfx86Updates + NumberOfFiles

        If _FileExists(UpdatesLocation$ + "\Answer_File\autounattend.xml") Then
            NumberOfx86Updates = NumberOfx86Updates + 1
            AddAnswerFile$ = "Y"
        Else
            AddAnswerFile$ = "N"
        End If

        If NumberOfx86Updates = 0 Then
            Cls
            Print
            Color 14, 4: Print "No x86 update files were found in this location.": Color 15
            Print "Please specify another location."
            If ScriptingChoice$ = "R" Then
                Print #5, ":: No x86 update files were found in this location."
                Print #5, ""
            End If
            Pause
            GoTo GetUpdatesLocation
        End If
    Case "DRIVERS", "BCD"
        FileTypeSearch x86Updates$ + "\", ".INF", "Y"
        NumberOfx86Updates = NumberOfFiles
        If NumberOfx86Updates = 0 Then
            Cls
            Print
            Color 14, 4: Print "No x86 drivers were found in this location.": Color 15
            Print "Please specify another location."
            If ScriptingChoice$ = "R" Then
                Print #5, ":: No x86 drivers were found in this location."
                Print #5, ""
            End If
            Pause
            GoTo GetUpdatesLocation
        End If
End Select

End_Getx86UpdatesLocation:

' Done with x86 updates location information

' Ask user what they want to name the final ISO image file

Cls
UserSelectedImageName$ = "" ' Set initial value
Print "If you would like to specify a name for the final ISO image file that this project will create, please do so now,"
Print "WITHOUT an extension. You can also simply press ENTER to use the default name of Windows.ISO."
Print
Print "Enter name ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension, or press ENTER: ";: Input "", UserSelectedImageName$

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
MkDir DestinationFolder$ + "WIM_x86"
MkDir DestinationFolder$ + "Assets"
MkDir DestinationFolder$ + "WinRE_MOUNT"
MkDir DestinationFolder$ + "WinPE_MOUNT"
MkDir DestinationFolder$ + "WinPE"
MkDir DestinationFolder$ + "WinRE"
MkDir DestinationFolder$ + "Setup_DU_x64"
MkDir DestinationFolder$ + "Setup_DU_x86"

' Export all the x64 and x86 editions to the WIM_x64 and WIM_x86 folders.

' Prior to starting the exports, we need to initialize some variables. To init these variables, we need to know how many indices we will be handling
' so we will determine that now.

TotalIndexCount = 0 'Init variable

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then
        TotalIndexCount = TotalIndexCount + IndexCount(x)
    End If
Next x

ReDim x64OriginalFile(0) As String
ReDim x86OriginalFile(0) As String
ReDim x64SourceArc(0) As String
ReDim x86SourceArc(0) As String
ReDim x64OriginalIndex(0) As String
ReDim x86OriginalIndex(0) As String
CurrentIndexCount = 0
x64ExportCount = 0
x86ExportCount = 0

' We are going to create a new PendingOps.log file. Delete old file if it exists.

If _FileExists(DestinationFolder$ + "logs\PendingOps.log") Then Kill DestinationFolder$ + "logs\PendingOps.log"

ff = FreeFile
Open (DestinationFolder$ + "logs\PendingOps.log") For Output As #ff
Print #ff, "Below is a list of files that were found with pending operations. Pending operations are the result of items added to the image that will"
Print #ff, "prevent DISM from being able to perform a cleanup operation on the image. The most common of the causes is enabling NetFX3. If this log"
Print #ff, "file lists any files below that have pending operations, you should redo this update using source files that do not have any pending"
Print #ff, "operations present."
Print #ff, ""
Print #ff, "NOTE: For any files listed below, the "; Chr$(34); "Architecture type"; Chr$(34); " will be shown as x64 or x86 if the original image file holds only x64 or x86"
Print #ff, "Windows Editions. If the line reads either "; Chr$(34); "x64_DUAL"; Chr$(34); " or "; Chr$(34); "x86_DUAL"; Chr$(34); ", then this edition of Windows is x64 or x86 as indicated, but it"
Print #ff, "comes from a source that is dual architecture."
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
            Select Case FileSourceType$(x)
                Case "x64"
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
                    + CHR$(34) + " BOOT.WIM /A-:RHS > NUL"
                    Shell _Hide Cmd$
                    Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\WinPE\BOOT.WIM" + Chr$(34) + " BOOT_x64.wim"
                    Shell _Hide Cmd$
                Case "x86"
                    x86ExportCount = x86ExportCount + 1
                    ReDim _Preserve x86OriginalFile(x86ExportCount) As String
                    ReDim _Preserve x86SourceArc(x86ExportCount) As String
                    ReDim _Preserve x86OriginalIndex(x86ExportCount) As String
                    x86OriginalFile$(x86ExportCount) = Temp$
                    x86SourceArc$(x86ExportCount) = "x86"
                    x86OriginalIndex$(x86ExportCount) = LTrim$(Str$(IndexList(x, y)))
                    SourceArcFlag$ = ""
                    DestArcFlag$ = "WIM_x86"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + SourceArcFlag$ + "\sources" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\WinPE"_
                    + CHR$(34) + " BOOT.WIM /A-:RHS > NUL"
                    Shell _Hide Cmd$
                    Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\WinPE\BOOT.WIM" + Chr$(34) + " BOOT_x86.wim"
                    Shell _Hide Cmd$
                Case "x64_DUAL"
                    x64ExportCount = x64ExportCount + 1
                    ReDim _Preserve x64OriginalFile(x64ExportCount) As String
                    ReDim _Preserve x64SourceArc(x64ExportCount) As String
                    ReDim _Preserve x64OriginalIndex(x64ExportCount) As String
                    x64OriginalFile$(x64ExportCount) = Temp$
                    x64SourceArc$(x64ExportCount) = "x64_DUAL"
                    x64OriginalIndex$(x64ExportCount) = LTrim$(Str$(IndexList(x, y)))
                    SourceArcFlag$ = "\x64"
                    DestArcFlag$ = "WIM_x64"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + SourceArcFlag$ + "\sources" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\WinPE"_
                    + CHR$(34) + " BOOT.WIM /A-:RHS > NUL"
                    Shell _Hide Cmd$
                    Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\WinPE\BOOT.WIM" + Chr$(34) + " BOOT_x64.wim"
                    Shell _Hide Cmd$
                Case "x86_DUAL"
                    x86ExportCount = x86ExportCount + 1
                    ReDim _Preserve x86OriginalFile(x86ExportCount) As String
                    ReDim _Preserve x86SourceArc(x86ExportCount) As String
                    ReDim _Preserve x86OriginalIndex(x86ExportCount) As String
                    x86OriginalFile$(x86ExportCount) = Temp$
                    x86SourceArc$(x86ExportCount) = "x86_DUAL"
                    x86OriginalIndex$(x86ExportCount) = LTrim$(Str$(IndexList(x, y)))
                    SourceArcFlag$ = "\x86"
                    DestArcFlag$ = "WIM_x86"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + SourceArcFlag$ + "\sources" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\WinPE"_
                    + CHR$(34) + " BOOT.WIM /A-:RHS > NUL"
                    Shell _Hide Cmd$
                    Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\WinPE\BOOT.WIM" + Chr$(34) + " BOOT_x86.wim"
                    Shell _Hide Cmd$
            End Select
            CurrentIndex$ = LTrim$(Str$(IndexList(x, y)))
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + MountedImageDriveLetter$ + SourceArcFlag$_
            + "\Sources\install.wim" + CHR$(34) + " /SourceIndex:" + CurrentIndex$ + " /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + DestArcFlag$_
            + "\install.wim" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Next y

        ' The next command dismounts the ISO image since we are now done with it. The messages displayed by the process are
        ' not really helpful so we are going to hide those messages even if detailed status is selected by the user.

        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    End If
Next x

' At this point, all images have been exported from their original files to the project folder \WIM_x64 and WIM_x86 folders.
' Now we need to mount each of those install.wim files and update all the images therein.

' Ditch the trailing backslash (robocopy does not like it)

CleanPath DestinationFolder$
DestinationFolder$ = Temp$
TotalImages = x64UpdateImageCount + x86UpdateImageCount

' The following section is run for either x64 or x86 editions that are having Windows updates injected

' We begin by updating the WinRE.wim and the Boot.wim (for WinPE). We only need to process the WinRE and boot.wim once
' so we will do that on the first x64 and first x86 edition that we process.

CurrentImage = 0

If (x64UpdateImageCount > 0 And InjectionMode$ = "UPDATES") Then
    For x = 1 To x64UpdateImageCount
        CurrentImage = CurrentImage + 1
        Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
        CurrentIndex$ = LTrim$(Str$(x))
        If x = 1 Then
            AddUpdatesStatusDisplay CurrentImage, TotalImages, 2
        Else
            AddUpdatesStatusDisplay CurrentImage, TotalImages, 25
        End If
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WIM_x64\install.wim" + CHR$(34)_
        + " /Index:" + CurrentIndex$ + " /Mountdir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$_
        + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Check the current Windows edition to see if it has any pending operations.

        Cmd$ = "PowerShell " + Chr$(34) + "Get-WindowsCapability -path '" + DestinationFolder$ + "\mount" + "' | Where-Object { $_.State -eq 'InstallPending' }" + Chr$(34) + " > capabilities.txt"
        Shell Cmd$
        ff = FreeFile
        Open "capabilities.txt" For Binary As #ff
        OpsPendingFileCheck$ = Space$(LOF(ff))
        Get #ff, 1, OpsPendingFileCheck$
        Close #ff

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

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 3
        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\mount\Windows\System32\Recovery" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\WINRE"_
        + CHR$(34) + " winre.wim /A-:RSH > NUL"
        Shell _Hide Cmd$

        ' Mount the WinRE Image if updates are available

        FileTypeSearch x64Updates$ + "\LCU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            SSU_Update_Avail$ = "Y"
        Else
            SSU_Update_Avail$ = "N"
        End If

        FileTypeSearch x64Updates$ + "\SafeOS_DU\", ".CAB", "N"

        If NumberOfFiles > 0 Then
            SafeOS_DU_Avail$ = "Y"
        Else
            SafeOS_DU_Avail$ = "N"
        End If

        If (SSU_Update_Avail$ = "N") And (SafeOS_DU_Avail$ = "N") Then GoTo Skip_WINRE_Update_x64

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\winre.wim" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Add SSU Update to WinRE.WIM

        If SSU_Update_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Add SafeOS DU to WinRE.WIM

        If SafeOS_DU_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\SafeOS_DU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Skip_WINRE_Update_x64:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\WINRE.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\WINRE_x64.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' del the temp file

        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\WinRE\winre.wim" + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 1

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 4

        ' Mount the WinPE Image - Index 1, if SSU or LCU updates exist

        FileTypeSearch x64Updates$ + "\LCU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            SSU_Updates_Avail$ = "Y"
        Else
            SSU_Updates_Avail$ = "N"
        End If

        FileTypeSearch x64Updates$ + "\LCU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            LCU_Updates_Avail$ = "Y"
        Else
            LCU_Updates_Avail$ = "N"
        End If

        If Skip_PE_Updates$ = "Y" Then GoTo Export_PE_Index1
        If (SSU_Updates_Avail$ = "N") And (LCU_Updates_Avail$ = "N") Then GoTo Export_PE_Index1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Add SSU Update to BOOT.WIM

        If SSU_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Add LCU Update to BOOT.WIM

        If LCU_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index1:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 2

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 5

        ' Generic files, such as scripts, that are being added to the boot.wim only need to be added to index 2. For that reason, we search for such
        ' files now, rather than when processing index 1.

        FileTypeSearch x64Updates$ + "\PE_Files\", "*", "N"
        If NumberOfFiles > 0 Then
            PE_Files_Avail$ = "Y"
        Else
            PE_Files_Avail$ = "N"
        End If

        ' Mount the WinPE Image - Index 2, if updates are available

        If (SSU_Updates_Avail$ = "N") And (LCU_Updates_Avail$ = "N") And (PE_Files_Avail$ = "N") Then GoTo Export_PE_Index2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /Index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Add SSU Update to BOOT.WIM

        If (SSU_Updates_Avail$ = "Y") And (Skip_PE_Updates$ = "N") Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Add LCU Update to BOOT.WIM

        If (LCU_Updates_Avail$ = "Y") And (Skip_PE_Updates$ = "N") Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
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
        + " /cleanup-image /StartComponentCleanup" + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index2:

        ' export index 2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /Bootable /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /SourceIndex:2 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Cmd$ = "move /y " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x64.wim" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\assets" + Chr$(34)
        Shell _Hide Cmd$
        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x64.wim" + Chr$(34)

        SkipWinPEx64:

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 6

        If SSU_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /PackagePath="_
            + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 7

        If LCU_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /PackagePath="_
            + CHR$(34) + x64Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        Cmd$ = "copy " + CHR$(34) + DestinationFolder$ + "\Assets\winre_x64.wim" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\MOUNT\Windows\System32\Recovery\WinRE.WIM" + CHR$(34)
        Shell _Hide Cmd$
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 8
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Cleanup-Image /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)_
        + " /StartComponentCleanup /ResetBase /ScratchDir:" + CHR$(34) + DestinationFolder$ + "\Temp" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 9

        ' Check to see if other updates are available

        FileTypeSearch x64Updates$ + "\Other\", ".MSU", "N"
        If NumberOfFiles > 0 Then
            Other_Updates_Avail$ = "Y"
        Else
            Other_Updates_Avail$ = "N"
        End If

        If Other_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x64Updates$ + "\Other" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Looking for and expanding the Setup Dynamic Update if it exists

        FileTypeSearch x64Updates$ + "\Setup_DU\", ".CAB", "N"

        If NumberOfFiles > 0 Then
            Setup_DU$ = TempArray$(1)
            Cmd$ = "expand " + Chr$(34) + Setup_DU$ + Chr$(34) + " -F:* " + Chr$(34) + DestinationFolder$ + "\Setup_DU_x64" + Chr$(34) + " > NUL"
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 10
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Get-Packages /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " > "_
        + CHR$(34) + DestinationFolder$ + "\Logs\x64_" + LTRIM$(STR$(x)) + "UpdateResults.txt" + CHR$(34) + ""
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 11
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\logs\dism.log" + Chr$(34) + " dism.log_x64_" + Right$(Str$(x), (Len(Str$(x)) - 1))
        Shell _Hide Cmd$
        FindDISMLogErrors_SingleFile (DestinationFolder$ + "\logs\dism.log_x64_" + Right$(Str$(x), (Len(Str$(x)) - 1))), DestinationFolder$ + "\logs"
        If DISM_Error_Found$ = "Y" Then
            ErrorsWereFound$ = "Y"
        End If
    Next x
End If

' Repeat the above for x86

If (x86UpdateImageCount > 0 And InjectionMode$ = "UPDATES") Then
    For x = 1 To x86UpdateImageCount
        CurrentImage = CurrentImage + 1
        Temp$ = DestinationFolder$ + "\WIM_x86\install.wim"
        CurrentIndex$ = LTrim$(Str$(x))
        If x = 1 Then
            AddUpdatesStatusDisplay CurrentImage, TotalImages, 2
        Else
            AddUpdatesStatusDisplay CurrentImage, TotalImages, 25
        End If

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WIM_x86\install.wim"_
        + CHR$(34) + " /Index:" + CurrentIndex$ + " /Mountdir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Check the current Windows edition to see if it has any pending operations.

        Cmd$ = "PowerShell " + Chr$(34) + "Get-WindowsCapability -path '" + DestinationFolder$ + "\mount" + "' | Where-Object { $_.State -eq 'InstallPending' }" + Chr$(34) + " > capabilities.txt"
        Shell Cmd$
        ff = FreeFile
        Open "capabilities.txt" For Binary As #ff
        OpsPendingFileCheck$ = Space$(LOF(ff))
        Get #ff, 1, OpsPendingFileCheck$
        Close #ff

        If InStr(OpsPendingFileCheck$, "InstallPending") Then
            OpsPending$ = "Y"
            ff2 = FreeFile
            Open (DestinationFolder$ + "\logs\PendingOps.log") For Append As #ff2
            Print #ff2, "File name: "; x86OriginalFile$(x)
            Print #ff2, "Architecture type: "; x86SourceArc$(x)
            Print #ff2, "Index number: "; x86OriginalIndex$(x)
            Print #ff2, ""
            Close #ff2
        End If

        Kill "capabilities.txt"

        ' Skip WinRE and WinPE if these have already been processed

        If x > 1 Then
            GoTo SkipWinPEx86
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 3
        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\mount\Windows\System32\Recovery" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\WINRE" + CHR$(34) + " winre.wim /A-:RSH > NUL"
        Shell _Hide Cmd$

        ' Mount the WinRE Image if updates are available

        FileTypeSearch x86Updates$ + "\LCU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            SSU_Update_Avail$ = "Y"
        Else
            SSU_Update_Avail$ = "N"
        End If

        FileTypeSearch x86Updates$ + "\SafeOS_DU\", ".CAB", "N"

        If NumberOfFiles > 0 Then
            SafeOS_DU_Avail$ = "Y"
        Else
            SafeOS_DU_Avail$ = "N"
        End If

        If (SSU_Update_Avail$ = "N") And (SafeOS_DU_Avail$ = "N") Then GoTo Skip_WINRE_Update_x86

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\winre.wim" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Add SSU Update to WinRE.WIM

        If SSU_Update_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /PackagePath=" + CHR$(34)_
            + x86Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Add SafeOS DU to WinRE.WIM

        If SafeOS_DU_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /PackagePath=" + CHR$(34)_
            + x86Updates$ + "\SafeOS_DU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /Commit" + " /LogPath="_
        + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Skip_WINRE_Update_x86:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\WINRE.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\WINRE_x86.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' del the temp file

        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\WinRE\winre.wim" + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 1

        ' Mount the WinPE Image - Index 1, if updates are available

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 4

        ' Mount the WinPE Image - Index 1, if SSU or LCU updates exist

        FileTypeSearch x86Updates$ + "\LCU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            SSU_Updates_Avail$ = "Y"
        Else
            SSU_Updates_Avail$ = "N"
        End If

        FileTypeSearch x86Updates$ + "\LCU\", ".MSU", "N"

        If NumberOfFiles > 0 Then
            LCU_Updates_Avail$ = "Y"
        Else
            LCU_Updates_Avail$ = "N"
        End If

        If Skip_PE_Updates$ = "Y" Then GoTo Export_PE_Index1_x86
        If (SSU_Updates_Avail$ = "N") And (LCU_Updates_Avail$ = "N") Then GoTo Export_PE_Index1_x86

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x86.WIM" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Add SSU Update to BOOT.WIM

        If (SSU_Updates_Avail$ = "Y") And (Skip_PE_Updates$ = "N") Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
            + " /PackagePath=" + CHR$(34) + x86Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Add LCU Update to BOOT.WIM

        If (LCU_Updates_Avail$ = "Y") And (Skip_PE_Updates$ = "N") Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /PackagePath="_
            + CHR$(34) + x86Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        ' dismount

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index1_x86:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x86.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x86.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 2

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 5

        ' Generic files, such as scripts, that are being added to the boot.wim only need to be added to index 2. For that reason, we search for such
        ' files now, rather than when processing index 1.

        FileTypeSearch x86Updates$ + "\PE_Files\", "*", "N"
        If NumberOfFiles > 0 Then
            PE_Files_Avail$ = "Y"
        Else
            PE_Files_Avail$ = "N"
        End If

        ' Mount the WinPE Image - Index 2, if updates are available

        If (SSU_Updates_Avail$ = "N") And (LCU_Updates_Avail$ = "N") And (PE_Files_Avail$ = "N") Then GoTo Export_PE_Index_x86

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x86.WIM" + CHR$(34)_
        + " /Index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Add SSU Update to BOOT.WIM

        If SSU_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /PackagePath="_
            + CHR$(34) + x86Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Add LCU Update to BOOT.WIM

        If LCU_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /PackagePath="_
            + CHR$(34) + x86Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Add generic files such as scripts to the boot.wim.

        FileTypeSearch x86Updates$ + "\PE_Files\", "*", "N"

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

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index_x86:

        ' export index 2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /Bootable /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x86.WIM" + CHR$(34)_
        + " /SourceIndex:2 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x86.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = "move /y " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x86.wim" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\assets" + Chr$(34)
        Shell _Hide Cmd$
        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x86.wim" + Chr$(34)

        ' The WinRE and WinPE components have been updated. We will now proceed with updating of the main OS (install.wim).

        SkipWinPEx86:

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 6

        If SSU_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /PackagePath="_
            + CHR$(34) + x86Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 7

        If LCU_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /PackagePath="_
            + CHR$(34) + x86Updates$ + "\LCU" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        Cmd$ = "copy " + CHR$(34) + DestinationFolder$ + "\Assets\winre_x86.wim" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\MOUNT\Windows\System32\Recovery\WinRE.WIM" + CHR$(34)
        Shell _Hide Cmd$
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 8
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Cleanup-Image /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)_
        + " /StartComponentCleanup /ResetBase /ScratchDir:" + CHR$(34) + DestinationFolder$ + "\Temp" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 9

        ' Check to see if other updates are available

        FileTypeSearch x86Updates$ + "\Other\", ".MSU", "N"
        If NumberOfFiles > 0 Then
            Other_Updates_Avail$ = "Y"
        Else
            Other_Updates_Avail$ = "N"
        End If

        If Other_Updates_Avail$ = "Y" Then
            Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Package /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /PackagePath="_
            + CHR$(34) + x86Updates$ + "\Other" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        ' Looking for and expanding the Setup Dynamic Update if it exists

        FileTypeSearch x86Updates$ + "\Setup_DU\", ".CAB", "N"

        If NumberOfFiles > 0 Then
            Setup_DU$ = TempArray$(1)
            Cmd$ = "expand " + Chr$(34) + Setup_DU$ + Chr$(34) + " -F:* " + Chr$(34) + DestinationFolder$ + "\Setup_DU_x86" + Chr$(34) + " > NUL"
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 10
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Get-Packages /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " > " + CHR$(34)_
        + DestinationFolder$ + "\Logs\x86_" + LTRIM$(STR$(x)) + "UpdateResults.txt" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 11
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\logs\dism.log" + Chr$(34) + " dism.log_x86_" + Right$(Str$(x), (Len(Str$(x)) - 1))
        Shell _Hide Cmd$
        FindDISMLogErrors_SingleFile (DestinationFolder$ + "\logs\dism.log_x86_" + Right$(Str$(x), (Len(Str$(x)) - 1))), DestinationFolder$ + "\logs"
        If DISM_Error_Found$ = "Y" Then
            ErrorsWereFound$ = "Y"
        End If
    Next x
End If

' We will now inject boot-critical drivers into the WinRE and WinPE images. We only need to process the WinRE and boot.wim once
' so we will do that on the first x64 and first x86 edition that we process.

If (x64UpdateImageCount > 0 And InjectionMode$ = "BCD") Then
    For x = 1 To x64UpdateImageCount
        CurrentImage = CurrentImage + 1
        Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
        CurrentIndex$ = LTrim$(Str$(x))
        If x = 1 Then
            AddUpdatesStatusDisplay CurrentImage, TotalImages, 51
        End If
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WIM_x64\install.wim" + CHR$(34)_
        + " /Index:" + CurrentIndex$ + " /Mountdir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$_
        + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        If x > 1 Then
            GoTo SkipWinPEx64_BCD
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 52
        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\mount\Windows\System32\Recovery" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\WINRE"_
        + CHR$(34) + " winre.wim /A-:RSH > NUL"
        Shell _Hide Cmd$

        ' Mount the WinRE Image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\winre.wim" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Driver /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) +_
         " /Driver:" + x64Updates$ + CHR$(34) + " /recurse" + "/LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Skip_WINRE_Update_x64_BCD:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\WINRE.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\WINRE_x64.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' del the temp file

        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\WinRE\winre.wim" + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 1

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 53

        ' Mount the WinPE Image - Index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Driver /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) +_
         " /Driver:" + x64Updates$ + CHR$(34) + " /recurse" + "/LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index1_BCD:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 2

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 54

        ' Mount the WinPE Image - Index 2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /Index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Driver /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) +_
         " /Driver:" + x64Updates$ + CHR$(34) + " /recurse" + "/LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34)_
        + " /cleanup-image /StartComponentCleanup" + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index2_BCD:

        ' export index 2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /Bootable /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x64.WIM" + CHR$(34)_
        + " /SourceIndex:2 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = "move /y " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x64.wim" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\assets" + Chr$(34)
        Shell _Hide Cmd$
        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x64.wim" + Chr$(34)

        ' The WinRE and WinPE components have been updated. We will now proceed with updating of the main OS (install.wim).

        SkipWinPEx64_BCD:

        Cmd$ = "copy " + CHR$(34) + DestinationFolder$ + "\Assets\winre_x64.wim" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\MOUNT\Windows\System32\Recovery\WinRE.WIM" + CHR$(34)
        Shell _Hide Cmd$
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 55
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Cleanup-Image /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)_
        + " /StartComponentCleanup /ResetBase /ScratchDir:" + CHR$(34) + DestinationFolder$ + "\Temp" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 56
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Get-Packages /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " > "_
        + CHR$(34) + DestinationFolder$ + "\Logs\x64_" + LTRIM$(STR$(x)) + "UpdateResults.txt" + CHR$(34) + ""
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Unmounting and saving the Windows edition

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 57
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\logs\dism.log" + Chr$(34) + " dism.log_x64_" + Right$(Str$(x), (Len(Str$(x)) - 1))
        Shell _Hide Cmd$
        FindDISMLogErrors_SingleFile (DestinationFolder$ + "\logs\dism.log_x64_" + Right$(Str$(x), (Len(Str$(x)) - 1))), DestinationFolder$ + "\logs"
        If DISM_Error_Found$ = "Y" Then
            ErrorsWereFound$ = "Y"
        End If
    Next x
End If

' Repeat the above for x86

If (x86UpdateImageCount > 0 And InjectionMode$ = "BCD") Then
    For x = 1 To x86UpdateImageCount
        CurrentImage = CurrentImage + 1
        Temp$ = DestinationFolder$ + "\WIM_x86\install.wim"
        CurrentIndex$ = LTrim$(Str$(x))

        If x = 1 Then
            AddUpdatesStatusDisplay CurrentImage, TotalImages, 51
        End If

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WIM_x86\install.wim"_
        + CHR$(34) + " /Index:" + CurrentIndex$ + " /Mountdir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)

        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Start Process of updating WinRE (WinRE.WIM)

        If x > 1 Then
            GoTo SkipWinPEx86_BCD
        End If

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 52
        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\mount\Windows\System32\Recovery" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\WINRE" + CHR$(34) + " winre.wim /A-:RSH > NUL"
        Shell _Hide Cmd$

        ' Mount the WinRE Image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\winre.wim" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Driver /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) +_
         " /Driver:" + x86Updates$ + CHR$(34) + " /recurse" + "/LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINRE_MOUNT" + CHR$(34) + " /Commit" + " /LogPath="_
        + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Skip_WINRE_Update_x86_BCD:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WinRE\WINRE.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\WINRE_x86.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' del the temp file

        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\WinRE\winre.wim" + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 1

        ' Mount the WinPE Image - Index 1, if updates are available

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 53

        ' Mount the WinPE Image - Index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x86.WIM" + CHR$(34)_
        + " /Index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Driver /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) +_
         " /Driver:" + x86Updates$ + CHR$(34) + " /recurse" + "/LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        ' dismount

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index1_x86_BCD:

        ' export index 1

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x86.WIM" + CHR$(34)_
        + " /SourceIndex:1 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x86.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' Update WinPE (BOOT.WIM)
        ' Index 2

        AddUpdatesStatusDisplay CurrentImage, TotalImages, 54

        ' Mount the WinPE Image - Index 2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x86.WIM" + CHR$(34)_
        + " /Index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Add-Driver /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) +_
         " /Driver:" + x86Updates$ + CHR$(34) + " /recurse" + "/LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' cleanup image

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /cleanup-image /StartComponentCleanup"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        ' dismount

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Export_PE_Index_x86_BCD:

        ' export index 2

        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Export-Image /Bootable /SourceImageFile:" + CHR$(34) + DestinationFolder$ + "\WINPE\BOOT_x86.WIM" + CHR$(34)_
        + " /SourceIndex:2 /DestinationImageFile:" + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x86.WIM" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

        Cmd$ = "move /y " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x86.wim" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\assets" + Chr$(34)
        Shell _Hide Cmd$
        Shell _Hide "del " + Chr$(34) + DestinationFolder$ + "\winpe\boot_x86.wim" + Chr$(34)

        ' The WinRE and WinPE components have been updated. We will now proceed with updating of the main OS (install.wim).

        SkipWinPEx86_BCD:

        Cmd$ = "copy " + CHR$(34) + DestinationFolder$ + "\Assets\winre_x86.wim" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\MOUNT\Windows\System32\Recovery\WinRE.WIM" + CHR$(34)
        Shell _Hide Cmd$
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 55
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Cleanup-Image /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34)_
        + " /StartComponentCleanup /ResetBase /ScratchDir:" + CHR$(34) + DestinationFolder$ + "\Temp" + CHR$(34) + " /LogPath=" + CHR$(34)_
        + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 56
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Get-Packages /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " > "_
        + CHR$(34) + DestinationFolder$ + "\Logs\x86_" + LTRIM$(STR$(x)) + "UpdateResults.txt" + CHR$(34) + ""
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 57
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\logs\dism.log" + Chr$(34) + " dism.log_x86_" + Right$(Str$(x), (Len(Str$(x)) - 1))
        Shell _Hide Cmd$
        FindDISMLogErrors_SingleFile (DestinationFolder$ + "\logs\dism.log_x86_" + Right$(Str$(x), (Len(Str$(x)) - 1))), DestinationFolder$ + "\logs"
        If DISM_Error_Found$ = "Y" Then
            ErrorsWereFound = "Y"
        End If
    Next x
End If

' The following section is run for either x64 or x86 editions that are having drivers injected

If (x64UpdateImageCount > 0 And InjectionMode$ = "DRIVERS") Then
    For x = 1 To x64UpdateImageCount
        CurrentImage = CurrentImage + 1
        Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
        CurrentIndex$ = LTrim$(Str$(x))
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 17
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WIM_x64\install.wim" + CHR$(34)_
        + " /Index:" + CurrentIndex$ + " /Mountdir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$_
        + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 18
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\Mount" + CHR$(34) + " /Add-Driver /Driver:" + CHR$(34)_
        + x64Updates$ + CHR$(34) + " /RECURSE /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 19
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Get-Drivers /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " > " + CHR$(34)_
        + DestinationFolder$ + "\Logs\x64_" + LTRIM$(STR$(x)) + "UpdateResults.txt" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 20
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\logs\dism.log" + Chr$(34) + " dism.log_x64_" + Right$(Str$(x), (Len(Str$(x)) - 1))
        Shell _Hide Cmd$
        FindDISMLogErrors_SingleFile (DestinationFolder$ + "\logs\dism.log_x64_" + Right$(Str$(x), (Len(Str$(x)) - 1))), DestinationFolder$ + "\logs"
        If DISM_Error_Found$ = "Y" Then
            ErrorsWereFound$ = "Y"
        End If
    Next x
End If

' Repeat the above for x86

If (x86UpdateImageCount > 0 And InjectionMode$ = "DRIVERS") Then
    For x = 1 To x86UpdateImageCount
        CurrentImage = CurrentImage + 1
        Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
        CurrentIndex$ = LTrim$(Str$(x))
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 17
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\WIM_x86\install.wim" + CHR$(34)_
        + " /Index:" + CurrentIndex$ + " /Mountdir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$_
        + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 18
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Image:" + CHR$(34) + DestinationFolder$ + "\Mount" + CHR$(34) + " /Add-Driver /Driver:" + CHR$(34)_
        + x86Updates$ + CHR$(34) + " /RECURSE /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 19
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Get-Drivers /Image:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " > " + CHR$(34)_
        + DestinationFolder$ + "\Logs\x86_" + LTRIM$(STR$(x)) + "UpdateResults.txt" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 20
        Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Unmount-Image /MountDir:" + CHR$(34) + DestinationFolder$ + "\mount" + CHR$(34) + " /Commit"_
        + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Cmd$ = "rename " + Chr$(34) + DestinationFolder$ + "\logs\dism.log" + Chr$(34) + " dism.log_x86_" + Right$(Str$(x), (Len(Str$(x)) - 1))
        Shell _Hide Cmd$
        FindDISMLogErrors_SingleFile (DestinationFolder$ + "\logs\dism.log_x86_" + Right$(Str$(x), (Len(Str$(x)) - 1))), DestinationFolder$ + "\logs"
        If DISM_Error_Found$ = "Y" Then
            ErrorsWereFound$ = "Y"
        End If
    Next x
End If

_ConsoleTitle "WIM Tools Version " + ProgramVersion$ + " by Hannes Sehestedt"

If ((x64UpdateImageCount > 0) And (x86UpdateImageCount > 0)) Then
    AllFilesAreSameArc = 0
    SingleImageTag$ = ""
Else
    AllFilesAreSameArc = 1
    If (x64UpdateImageCount > 0) Then
        SingleImageTag$ = "\x64"
    ElseIf (x86UpdateImageCount > 0) Then
        SingleImageTag$ = "\x86"
    End If
End If

' To ensure that DestinationFolder$ is always specified consistently without a trailing backslash, we will
' run it through the CleanPath routine.

CleanPath DestinationFolder$
DestinationFolder$ = Temp$

' Jump to the routine for creating a base image. Here we determine whether we need to run code for a Single or Dual Architecture project

Select Case ProjectType$
    Case "x64", "x86"
        GoTo ProjectIsSingleArchitecture
    Case "DUAL"
        GoTo ProjectIsDualArchitecture
End Select

ProjectIsSingleArchitecture:

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then
        Temp$ = FolderArray$(x) + FileArray$(x)

        Select Case InjectionMode$
            Case "UPDATES"
                AddUpdatesStatusDisplay CurrentImage, TotalImages, 12
            Case "DRIVERS"
                AddUpdatesStatusDisplay CurrentImage, TotalImages, 21
            Case "BCD"
                AddUpdatesStatusDisplay CurrentImage, TotalImages, 58
        End Select

        MountISO Temp$

        ' If an x64 folder exists, then even though the project is a single architecture type project, the source is a dual architecture source.
        ' This means that we need to copy the contents of the x64 or x86 folder to the root and not to the x64 or x86 folder.

        If _DirExists(MountedImageDriveLetter$ + "\x64") Then
            Select Case ExcludeAutounattend$
                Case "Y"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + SingleImageTag$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
                    + CHR$(34) + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd autounattend.xml > NUL"
                Case "N"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + SingleImageTag$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
                    + CHR$(34) + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd > NUL"
            End Select
            Shell _Hide Cmd$
        Else
            Select Case ExcludeAutounattend$
                Case "Y"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
                    + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd autounattend.xml > NUL"
                Case "N"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
                    + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd > NUL"
            End Select
            Shell _Hide Cmd$
        End If
        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Exit For
    End If
Next x

' When we arrive here, the base image for the single architecure type project has been completed.

GoTo DoneCreatingBaseImage

ProjectIsDualArchitecture:

Select Case InjectionMode$
    Case "UPDATES"
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 12
    Case "DRIVERS"
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 21
    Case "BCD"
        AddUpdatesStatusDisplay CurrentImage, TotalImages, 58
End Select

' If DualArcImagePath$ is not empty then we have a dual architecture image in this project and we can use it to build the
' base image. Otherwise, we will need to use files from both the x64 and x86 distributions and we need to dynamically
' create a pair of "bcd" files.

If DualArcImagePath$ = "" Then GoTo NoDualImageDistrib

MountISO DualArcImagePath$

Select Case ExcludeAutounattend$
    Case "Y"
        Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + " " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
        + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd autounattend.xml > NUL"
    Case "N"
        Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + " " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
        + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd > NUL"
End Select
Shell _Hide Cmd$
Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + DualArcImagePath$ + "'" + Chr$(34) + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

NoDualImageDistrib:

' If no dual architecture image is available, copy the needed files from x64 and x86 media
' and dynamically generate the bcd files needed.

For x = 1 To TotalFiles
    If FileSourceType$(x) = "x64" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        MountISO Temp$
        Select Case ExcludeAutounattend$
            Case "Y"
                Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xd sources support /xf bcd install.wim autounattend.xml /a-:rsh > NUL"
            Case "N"
                Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xd sources support /xf bcd install.wim /a-:rsh > NUL"
        End Select
        Shell _Hide Cmd$
        Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x64" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xf install.wim autounattend.xml /a-:rsh > NUL"
        Shell _Hide Cmd$
        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-Diskimage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        GoTo Do_x86_Build_Base
    End If
Next x

Do_x86_Build_Base:

For x = 1 To TotalFiles
    If FileSourceType$(x) = "x86" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        MountISO Temp$
        Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " autorun.inf setup.exe /a-:rsh > NUL"
        Shell _Hide Cmd$
        Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + "\efi\boot" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\efi\boot" + Chr$(34) + " bootia32.efi /a-:rsh > NUL"
        Shell _Hide Cmd$
        Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x86" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xf install.wim autounattend.xml /a-:rsh > NUL"
        Shell _Hide Cmd$
        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-Diskimage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        GoTo Build_Bcd_Files
    End If
Next x

Build_Bcd_Files:

' Create .reg files

GoSub Create_Reg_Files

' Create a template bcd hive for each of the two locations where we need this file

Cmd$ = "bcdedit /createstore " + Chr$(34) + DestinationFolder$ + "\ISO_Files\boot\bcd" + Chr$(34)
Shell _Hide Cmd$
Cmd$ = "bcdedit /createstore " + Chr$(34) + DestinationFolder$ + "\ISO_Files\efi\microsoft\boot\bcd" + Chr$(34)
Shell _Hide Cmd$

' Load the template hives into the registry

Cmd$ = "reg load HKLM\BCD_BIOS " + Chr$(34) + DestinationFolder$ + "\ISO_Files\boot\bcd" + Chr$(34)
Shell _Hide Cmd$
Cmd$ = "reg load HKLM\BCD_EFI " + Chr$(34) + DestinationFolder$ + "\ISO_Files\efi\microsoft\boot\bcd" + Chr$(34)
Shell _Hide Cmd$

' Create a "permissions.txt" file that defines registry permission changes to be applied by regini

Temp$ = TempLocation$ + "\permissions.txt"
Open Temp$ For Output As #1
Print #1, "\Registry\machine\BCD_BIOS [1 6 17]"
Print #1, "\Registry\machine\BCD_BIOS\Description [1 6 17]"
Print #1, "\Registry\machine\BCD_BIOS\Objects [1 6 17]"
Print #1, "\Registry\machine\BCD_EFI [1 6 17]"
Print #1, "\Registry\machine\BCD_EFI\Description [1 6 17]"
Print #1, "\Registry\machine\BCD_EFI\Objects [1 6 17]"
Close #1

' Run regini to alter the permissions

Cmd$ = "regini " + Chr$(34) + TempLocation$ + "\permissions.txt" + Chr$(34)
Shell _Hide Cmd$
Temp$ = TempLocation$ + "\permissions.txt"
Kill Temp$

' Import the registry files in order to apply those settings to out template hives

Cmd$ = "reg import " + Chr$(34) + TempLocation$ + "\bcd_bios.reg" + Chr$(34)
Shell _Hide Cmd$
Cmd$ = "reg import " + Chr$(34) + TempLocation$ + "\bcd_efi.reg" + Chr$(34)
Shell _Hide Cmd$

' Unload the registry hives, committing the changes to what were the templates, making them
' the final bcd hive files.

Cmd$ = "reg unload HKLM\BCD_BIOS"
Shell _Hide Cmd$
Cmd$ = "reg unload HKLM\BCD_EFI"
Shell _Hide Cmd$

' Delete the temporary registry files used to create the bcd files

Temp$ = TempLocation$ + "\bcd_bios.reg"
Kill Temp$
Temp$ = TempLocation$ + "\bcd_efi.reg"
Kill Temp$

' Done with creation of dual architure base image

DoneCreatingBaseImage:

' When we arrive here, the creation of the dual architecture base image is completed

' Create an ei.cfg file within the base image if the user chose to do so

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

    Temp$ = DestinationFolder$ + "\ISO_Files\x64\sources"

    If _DirExists(Temp$) Then
        If Not (_FileExists(Temp$ + "\ei.cfg")) Then
            Open (Temp$ + "\ei.cfg") For Output As #1
            Print #1, "[CHANNEL]"
            Print #1, "Retail"
            Close #1
        End If
    End If

    Temp$ = DestinationFolder$ + "\ISO_Files\x86\sources"

    If _DirExists(Temp$) Then
        If Not (_FileExists(Temp$ + "\ei.cfg")) Then
            Open (Temp$ + "\ei.cfg") For Output As #1
            Print #1, "[CHANNEL]"
            Print #1, "Retail"
            Close #1
        End If
    End If
End If

' Done processing ei.cfg

' Moving the updated install.wim, Boot.wim, and Setup Dynamic Updates to the base image

Select Case InjectionMode$
    Case "UPDATES"
        AddUpdatesStatusDisplay 0, 0, 13
    Case "DRIVERS"
        AddUpdatesStatusDisplay 0, 0, 22
    Case "BCD"
        AddUpdatesStatusDisplay 0, 0, 59
End Select

If ProjectType$ = "DUAL" Then
    If AllFilesAreSameArc = 1 Then
Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\WIM_x64\install.wim " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
        + "\Sources" + CHR$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\WIM_x86\install.wim " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
        + "\Sources" + CHR$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
        + "\Sources\BOOT.WIM" + CHR$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x86.WIM " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
        + "\Sources\BOOT.WIM" + CHR$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\Setup_DU_x64" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + "\Sources"_
        + CHR$(34) + " *.* /e > NUL"
        Shell _Hide Cmd$

        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\Setup_DU_x86" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
        + "\Sources" + CHR$(34) + " *.* /e > NUL"
        Shell _Hide Cmd$
    Else
        Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "\WIM_x64\install.wim " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x64\Sources" + Chr$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "\WIM_x86\install.wim " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x86\Sources" + Chr$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM " + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\ISO_Files\x64\Sources\BOOT.WIM" + CHR$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x86.WIM " + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
        + "\ISO_Files\x86\Sources\BOOT.WIM" + CHR$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "robocopy " + Chr$(34) + DestinationFolder$ + "\Setup_DU_x64" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x64\Sources" + Chr$(34) + " *.* /e > NUL"
        Shell _Hide Cmd$
        Cmd$ = "robocopy " + Chr$(34) + DestinationFolder$ + "\Setup_DU_x86" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x86\Sources" + Chr$(34) + " *.* /e > NUL"
        Shell _Hide Cmd$
    End If
End If

If ProjectType$ = "SINGLE" Then
    If x64UpdateImageCount > 0 Then
        Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
        If _FileExists(Temp$) Then
            Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
            Shell _Hide Cmd$
            Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "\WIM_x64\install.wim " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
            Shell _Hide Cmd$
            Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x64.WIM " + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
            + "\ISO_Files\Sources\BOOT.WIM" + CHR$(34) + " > NUL"
            Shell _Hide Cmd$
                        Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\Setup_DU_x64" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
                       + "\ISO_Files\Sources" + CHR$(34) + " *.* /e > NUL"
            Shell _Hide Cmd$
        End If
    End If
    If x86UpdateImageCount > 0 Then
        Temp$ = DestinationFolder$ + "\WIM_x86\install.wim"
        If _FileExists(Temp$) Then
            Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
            Shell _Hide Cmd$
            Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "\WIM_x86\install.wim " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
            Shell _Hide Cmd$
            Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\Assets\BOOT_x86.WIM " + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
            + "\ISO_Files\Sources\BOOT.WIM" + CHR$(34) + " > NUL"
            Shell _Hide Cmd$
            Cmd$ = "robocopy " + CHR$(34) + DestinationFolder$ + "\Setup_DU_x86" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
            + "\ISO_Files\Sources" + CHR$(34) + " *.* /e > NUL"
            Shell _Hide Cmd$
        End If
    End If
End If

' This code fixes a problem with the Windows update procedure. When WinPE (boot.wim) index #2 is updated, it's entirely
' possible that some duplicate files in other locations may not get updated. Those duplicate files are supposed to be
' exactly the same as the ones updated in the boot.wim. This procedure will assure that this is the case.

' If a \Sources folder exists right off of the root of the base image, then we have a single architecture.

If _DirExists(DestinationFolder$ + "\ISO_Files\Sources") Then

Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\ISO_Files\Sources\boot.wim" + CHR$(34)_
 + " /index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "copy /B " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\setup.exe" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /Y"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "copy /B " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\Sources\setup.exe" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " /Y"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "robocopy " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\Windows\System32" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " *.* /e /ndl /xo /xx /xl /np /r:0 /w:0 > NUL"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /Mountdir:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /discard /LogPath=" + Chr$(34) + DestinationFolder$ + "\Logs\dism.log" + Chr$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\ISO_Files\Sources\install.wim" + CHR$(34)_
 + " /index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "robocopy " + Chr$(34) + DestinationFolder$ + "\MOUNT\Windows\System32" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " *.* /e /ndl /xo /xx /xl /np /r:0 /w:0 > NUL"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /Mountdir:" + Chr$(34) + DestinationFolder$ + "\MOUNT" + Chr$(34) + " /discard /LogPath=" + Chr$(34) + DestinationFolder$ + "\Logs\dism.log" + Chr$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

Else

    ' We reach this if the project is not single architecture. We need to hanle both the x64 and x86 elements.

    ' Handle x64

Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\ISO_Files\x64\Sources\boot.wim" + CHR$(34)_
 + " /index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "copy /B " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\setup.exe" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x64" + Chr$(34) + " /Y"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "copy /B " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\Sources\setup.exe" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x64\Sources" + Chr$(34) + " /Y"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "robocopy " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\Windows\System32" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x64\Sources" + Chr$(34) + " *.* /e /ndl /xo /xx /xl /np /r:0 /w:0 > NUL"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /Mountdir:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /discard /LogPath=" + Chr$(34) + DestinationFolder$ + "\Logs\dism.log" + Chr$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\ISO_Files\x64\Sources\install.wim" + CHR$(34)_
 + " /index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "robocopy " + Chr$(34) + DestinationFolder$ + "\MOUNT\Windows\System32" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x64\Sources" + Chr$(34) + " *.* /e /ndl /xo /xx /xl /np /r:0 /w:0 > NUL"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /Mountdir:" + Chr$(34) + DestinationFolder$ + "\MOUNT" + Chr$(34) + " /discard /LogPath=" + Chr$(34) + DestinationFolder$ + "\Logs\dism.log" + Chr$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

    ' Handle x86

Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\ISO_Files\x86\Sources\boot.wim" + CHR$(34)_
 + " /index:2 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\WINPE_MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "copy /B " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\setup.exe" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x86" + Chr$(34) + " /Y"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

    ' For a dual architecture project, there is one additional setup.exe file that we need to deal with. The root of the main media has a setup that should
    ' be the same as the \x86\setup.exe

    Cmd$ = "copy /B " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\setup.exe" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /Y"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "copy /B " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\Sources\setup.exe" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x86\Sources" + Chr$(34) + " /Y"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "robocopy " + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT\Windows\System32" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x86\Sources" + Chr$(34) + " *.* /e /ndl /xo /xx /xl /np /r:0 /w:0 > NUL"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /Mountdir:" + Chr$(34) + DestinationFolder$ + "\WINPE_MOUNT" + Chr$(34) + " /discard /LogPath=" + Chr$(34) + DestinationFolder$ + "\Logs\dism.log" + Chr$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

Cmd$ = CHR$(34) + DISMLocation$ + CHR$(34) + " /Mount-Image /ImageFile:" + CHR$(34) + DestinationFolder$ + "\ISO_Files\x86\Sources\install.wim" + CHR$(34)_
 + " /index:1 /Mountdir:" + CHR$(34) + DestinationFolder$ + "\MOUNT" + CHR$(34) + " /LogPath=" + CHR$(34) + DestinationFolder$ + "\Logs\dism.log" + CHR$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = "robocopy " + Chr$(34) + DestinationFolder$ + "\MOUNT\Windows\System32" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x86\Sources" + Chr$(34) + " *.* /e /ndl /xo /xx /xl /np /r:0 /w:0 > NUL"
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Unmount-Image /Mountdir:" + Chr$(34) + DestinationFolder$ + "\MOUNT" + Chr$(34) + " /discard /LogPath=" + Chr$(34) + DestinationFolder$ + "\Logs\dism.log" + Chr$(34)
    Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

End If

' End of workaraound for Microsoft update procedure issue.

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
    AnswerFilePresent = "Y"
Else AnswerFilePresent = "N"
End If

Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -m -o -u2 -udfver102 -bootdata:2#p0,e,b" + CHR$(34) + DestinationFolder$_
 + "\ISO_Files\boot\etfsboot.com" + CHR$(34) + "#pEF,e,b" + CHR$(34) + DestinationFolder$ + "\ISO_Files\efi\microsoft\boot\efisys.bin"_
 + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34) + " " + CHR$(34) + FinalImageName$ + CHR$(34) +_
 " >> " + CHR$(34) + DestinationFolder$ + "\logs\OSCDIMG.log" + CHR$(34) + " 2>&1"
ff = FreeFile
Open DestinationFolder$ + "\logs\OSCDIMG.log" For Output As #ff
Print #ff, Cmd$
Print #ff, ""
Close #ff
Shell Chr$(34) + Cmd$ + Chr$(34)

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
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\WIM_x86" + Chr$(34) + " /s /q"
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
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\Setup_DU_x86" + Chr$(34) + " /s /q"
Shell _Hide Cmd$

If ErrorsWereFound$ = "Y" Then

    ' The lines of code below are used to create a new log file as a result of a known Microsoft issue that is causing a number of
    ' false error messages to be logged. As a result of this, we parse throgh all the errors that are being saved in the file
    ' ERROR_SUMMARY.log and strip off the known false errors. We then stuff the results into the file named SANITIZED_ERROR_SUMMARY.log.

    ErrMsgFile1 = FreeFile
    Open DestinationFolder$ + "\logs\ERROR_SUMMARY.log" For Input As #ErrMsgFile1
    ErrMsgFile2 = FreeFile
    Open DestinationFolder$ + "\logs\SANTIZED_ERROR_SUMMARY.log" For Output As #ErrMsgFile2
    Do Until EOF(ErrMsgFile1)
        Line Input #ErrMsgFile1, ErrMsg$(1)
        If Left$(ErrMsg$(1), 8) = "Warning!" Then
            For x = 2 To 4
                Line Input #ErrMsgFile1, ErrMsg$(x)
            Next x

            ' Messages to strip from log

            If InStr(ErrMsg$(4), "Matching binary") <> 0 And InStr(ErrMsg$(4), "missing for component") <> 0 Then
                GoTo NoPrintSanitized
            End If

            If InStr(ErrMsg$(4), "Error 800f0984 [Warning,Facility=15 (0x000f),Code=2436 (0x0984)] originated in function ComponentStore") Then
                GoTo NoPrintSanitized
            End If

            If InStr(ErrMsg$(4), "Registry overridable collision found") Then
                GoTo NoPrintSanitized
            End If

            If InStr(ErrMsg$(4), "One of the components setting this value is Windows-Defender-Service") Then
                GoTo NoPrintSanitized
            End If

            If InStr(ErrMsg$(4), "Previously seen component setting this value is Windows-Defender-Service") Then
                GoTo NoPrintSanitized
            End If

            If InStr(ErrMsg$(4), "(hr:0xc1420117)") Then
                GoTo NoPrintSanitized
            End If

            If InStr(ErrMsg$(4), "One of the components setting this value is Microsoft-Windows-Authentication-AuthUI-Component") Then
                GoTo NoPrintSanitized
            End If

            If InStr(ErrMsg$(4), "Previously seen component setting this value is Microsoft-Windows-Authentication-AuthUI-Component") Then
                GoTo NoPrintSanitized
            End If

            If InStr(ErrMsg$(4), "One of the components setting this value is Microsoft-Windows-Security-SPP") Then
                GoTo NoPrintSanitized
            End If

            If InStr(ErrMsg$(4), "Previously seen component setting this value is Microsoft-Windows-Security-SPP") Then
                GoTo NoPrintSanitized
            End If

            ' End of messages to strip from log

            For x = 1 To 4
                Print #ErrMsgFile2, ErrMsg$(x)
            Next x
            Print #ErrMsgFile2, ""

            NoPrintSanitized:

        End If

    Loop
    Close #ErrMsgFile2
    Close #ErrMsgFile1

End If

' Remove the AV exclusion for the destination folder

CleanPath DestinationFolder$
Cmd$ = "powershell.exe -command Remove-MpPreference -ExclusionPath " + "'" + Chr$(34) + Temp$ + Chr$(34) + "'"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
If _FileExists("WIM_Exclude_Path.txt") Then Kill "WIM_Exclude_Path.txt"

' If the user opted to have a shutdown performed, then skip to the "ShutdownRequested" routine.

If _FileExists(Environ$("userprofile") + "\Desktop\Auto_Shutdown.txt") Then GoTo shutdownrequested

NoShutdown:

Cls
Print
Color 0, 10: Print "All processes have been completed.": Color 15

If _FileExists(DestinationFolder$ + "\logs\SANITIZED_ERROR_SUMMARY.log") Then
    Print
    Print "It is suggested that you review the contents of the log file named ";: Color 10: Print "SANITIZED_ERROR_SUMMARY.log";: Color 15: Print " to make certain"
    Print "that there were no unexpected errors."
    Print
    Print "You can find the log file here:"
    Print
    Color 10: Print DestinationFolder$; "\logs\SANITIZED_ERROR_SUMMARY.log": Color 15
End If

Pause

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

shutdownrequested:

' User elected to have an automatic shutdown performed. We will give then one last opportunity to abort.

Cls
Print
Color 0, 10: Print "All processes have been completed.": Color 15
Print
Print "An automatic shutdown of the system has been requested. You can abort the shutdown if you delete or rename the file"

For x = 60 To 0 Step -1
    _Limit 1
    Locate 5, 1: Print "named ";: Color 10: Print "AUTO_SHUTDOWN.txt";: Color 15: Print " located on your desktop within the next";: Color 10: Print x;: Color 15: Print "seconds.  "
    If Not (_FileExists(Environ$("userprofile") + "\Desktop\Auto_Shutdown.txt")) Then GoTo NoShutdown
Next x

ff = FreeFile
Open ("WIM_Shutdown_log.txt") For Output As #ff
Print #ff, "Program run was completed on "; Date$; " at "; Time$

If _FileExists(DestinationFolder$ + "\logs\SANITIZED_ERROR_SUMMARY.log") Then
    Print #ff, ""
    Print #ff, "It is suggested that you review the contents of the log file named SANITIZED_ERROR_SUMMARY.log to make certain"
    Print #ff, "that there were no unexpected errors."
    Print #ff, ""
    Print #ff, "You can find the log file here:"
    Print #ff, ""
    Print #ff, DestinationFolder$; "\logs\SANITIZED_ERROR_SUMMARY.log"
End If

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
Shell _Hide "shutdown /s /t 5 /f"
GoTo EndProgram


' **************************************************************************************************
' * Make or update a bootable drive from one or more Windows ISO images and Windows PE / RE images *
' **************************************************************************************************

MakeBootDisk:

' This routine will allow you to create a bootable drive for installing Windows or
' to be used as an emergency boot disk. The option to create one or more additional
' generic partitions is also provided as is support for BitLocker in the single boot
' image option.

' IMPORTANT: This routine has some code to allow the user to choose whether any partitions aside from the
' first partition should be created as NTFS or exFAT. If you want the program to just automatically
' create all partitions as NTFS then set "UserCanPickFS$" below to FALSE. It you set it to TRUE then
' the user will be allowed to pick the filesystem to use.

UserCanPickFS$ = "FALSE" ' Set this to "FALSE" to always use NTFS. See the note above about this.

Boot_ModeSelect:

Do
    Cls
    Print "Do you want to create a Single bootable partition or Multple boot partitions?"
    Print
    Print "If you are not yet familiar with this option, enter HELP below for important things to know about this option."
    Print
    Input "Single, Multiple, or HELP: "; Temp$
    Temp$ = UCase$(Temp$)
    Select Case Temp$

        Case "S", "M", "H", "SINGLE", "MULTIPLE", "HELP"
            Temp$ = Left$(Temp$, 1)
            Exit Do
    End Select
Loop

If Temp$ = "H" Then
    Cls
    Print "If you wish to take a single Window ISO image and make bootable media from it, choose the Single option. Note that"
    Print "this option will still give you the ability to create up to two additional partitions that you can use for storing"
    Print "other data. This mode is highly compatible with x86, x64, BIOS, and UEFI based systems. This method uses an MBR"
    Print "(Master Boot Record) disk configuration whereas the other method uses a GPT (GUID Partition Table) configuration."
    Print
    Print "The Multiple option will sacrifice x86 and BIOS compatibility but offers you the ability to do the following:"
    Print
    Print "1) You can boot multiple different items. For example, you could boot your Window 10 media, Windows 11 media, Macrium"
    Print "   reflect recovery media, and more all from the same flash drive, SSD, HDD, etc."
    Print
    Print "2) You will leave behind the 2TB disk limitation. You can use a disk of any size."
    Print
    Print "3) Rather than a 4 partition limit, this program will support up to 15 partitions"
    Print
    Print "This option is intended only for use with x64 UEFI based systems."
    Pause
    GoTo Boot_ModeSelect
End If

If Temp$ = "S" Then GoTo Boot_SingleMode
If Temp$ = "M" Then GoTo Boot_MultipleMode

Boot_SingleMode:

AddPart$ = ""
TotalPartitions = 0
AdditionalPartitions = 0
ReDim PartitionSize(4) As String
ReDim BitLockerFlag(4) As String
ReDim AutoUnlock(4) As String
ReDim Letter(4) As String
ReDim VolLabel(4) As String

VolLabel(1) = "Partition 1"
VolLabel(2) = "Partition 2"
VolLabel(3) = "Partition 3"
VolLabel(4) = "Partition 4"

' Get Windows ISO path to copy to the thumb drive

GetSourceISOForMakeBoot:

MakeBootableSourceISO$ = "" ' Set initial value

Do
    Cls
    Print "Enter the full path including the file name for the Windows ISO image you want to copy to the drive."
    Print
    Input "Enter the full path: ", MakeBootableSourceISO$
Loop While MakeBootableSourceISO$ = ""

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
    Print " This will completely wipe the contents of a disk and configure it from scratch."
    Print
    Color 0, 10: Print "2)";: Color 15: Print " ";: Color 0, 14: Print "REFRESH DISK:";: Color 15
    Print " This will leave all partitions and data intact except the Windows installation partitions which will"
    Print "   be updated and refreshed to match the ISO image you just specified. This option is intended only for disks"
    Print "   previously initialized by this routine where a WIPE operation was performed."
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

If WipeOrRefresh = 1 GoTo AskForPartitions

' We arrive here if the user elected to perform a refresh

GetDriveInfo_FAT32:

' Since we are performing a refresh operation, the drive letters for the Windows installation media should already be present.
' We will now search for these drive letters. In addition, we need to make sure that there is not more than one Windows
' installation disk connected to the system to make sure that we do not accidentally select the wrong disk.

' Initialize variables

Par1InstancesFound = 0
Par2InstancesFound = 0

Restore DriveLetterData

For x = 1 To 24
    Read MediaLetter$
    If _FileExists(MediaLetter$ + ":\PAR1_MEDIA.WIM") Then
        Par1InstancesFound = Par1InstancesFound + 1
    End If
Next x

Restore DriveLetterData

For x = 1 To 24
    Read MediaLetter$
    If _FileExists(MediaLetter$ + ":\PAR2_MEDIA.WIM") Then
        Par2InstancesFound = Par2InstancesFound + 1
    End If
Next x

If Par1InstancesFound = 1 And Par2InstancesFound = 1 Then
    GoTo DetectBootMediaLetters
Else
    GoTo GetBootMediaDriveLetters
End If

DetectBootMediaLetters:

' If we get here, then we already know that there is boot media previously created by this program connected to the system.
' We also know that there is only one such media connected. We will scan all the drives to determine what the drive letter
' for each of the first two partitions are.

Restore DriveLetterData

For x = 1 To 24
    Read MediaLetter$
    If _FileExists(MediaLetter$ + ":\PAR1_MEDIA.WIM") Then
        FAT32DriveLetter$ = MediaLetter$
        Exit For
    End If
Next x

Restore DriveLetterData

For x = 1 To 24
    Read MediaLetter$
    If _FileExists(MediaLetter$ + ":\PAR2_MEDIA.WIM") Then
        exFATorNTFSdriveletter$ = MediaLetter$
        Exit For
    End If
Next x

GetBootMediaDriveLetters:

' If Par1InstancesFound is equal 1 then we already have a valid drive letter so we will jump to the point where we do a check to make
' sure that the drive contains valid data.

If Par1InstancesFound = 1 Then GoTo TestPar1

Cls
Print "We need to know the drive letter of the partitions used to install Windows."
Print
Print "Please enter the ";: Color 0, 14: Print "FAT32";: Color 15: Print " partition drive letter (Partition #1). Enter only the letter (no colon): ";: Input "", FAT32DriveLetter$

If Len(FAT32DriveLetter$) <> 1 Then
    Cls
    Color 14, 4: Print "Invalid response!": Color 15
    Print
    Print "Please enter a drive letter only!"
    Pause
    GoTo GetDriveInfo_FAT32
End If

TestPar1:

Temp1$ = FAT32DriveLetter$ + ":\boot"
Temp2$ = FAT32DriveLetter$ + ":\efi"

If Not ((_DirExists(Temp1$)) Or (_DirExists(Temp2$))) Then
    Cls
    Color 14, 4: Print "Warning!": Color 15
    Print
    Print "This does not seem to be a valid FAT32 boot partition. As a precation we check for the existance of certain"
    Print "folders which we have not found. Please check to make sure you have chosen the correct drive."
    Pause
    GoTo GetDriveInfo_FAT32
End If

GetDriveInfo_exFATorNTFS:

' If Par2InstancesFound is equal 1 then we already have a valid drive letter so we will jump to the point where we do a check to make
' sure that the drive contains valid data.

If Par2InstancesFound = 1 Then GoTo TestPar2

Cls
Print "Now, enter the second partition ";: Color 0, 14: Print "("; FSType$; ")";: Color 15: Print " drive letter. Enter only the letter (no colon): ";: Input "", exFATorNTFSdriveletter$

If Len(exFATorNTFSdriveletter$) <> 1 Then
    Cls
    Color 14, 4: Print "Invalid response!": Color 15
    Print
    Print "Please enter a drive letter only!"
    Pause
    GoTo GetDriveInfo_exFATorNTFS
End If

TestPar2:

Temp1$ = exFATorNTFSdriveletter$ + ":\x64\sources\install.wim"
Temp2$ = exFATorNTFSdriveletter$ + ":\sources\install.wim"

If Not ((_FileExists(Temp1$)) Or (_FileExists(Temp2$))) Then
    Cls
    Color 14, 4: Print "Warning!": Color 15
    Print
    Print "This does not seem to be a valid install partition. As a precation we check for the existance of certain"
    Print "folders which we have not found. Please check to make sure you have chosen the correct drive."
    Pause
    GoTo GetDriveInfo_exFATorNTFS
End If

' When we arrive, then we have what appears to be valid drive letters. We will now mount the ISO image so that we can
' refresh the thumb drive with the contents.

' We already know the architecture of the ISO image being used for the refresh from the earlier mount of the image
' so there is no need to check it again. Architecture = 1 if it is a single architecture image, 2 if dual architecture
' We will now clean off the thumb drive partitions and copy the refreshed data to it. This gives us everything we need
' to jump to the already existing existing routine that mounts the ISO Image and copies file to the drive EXCEPT
' that this routine is expecting the drive letter of the FAT32 partition in Letter$(1) and the exFAT / NTFS partition
' in Letter$(2).

' We will also format those partitions first to clar the current content

' We want to keep the exiting volume labels, but these are destroyed when we format the partitions. We will save the current volume labels
' and then restore them after the format.

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

Cls
Print "We are now preparing the partitions."

' Format partition 1

Cmd$ = "format " + FAT32DriveLetter$ + ": /FS:FAT32 /Q /Y > NUL"
Shell Cmd$

' Restore the volume label if the original volume label was not blank.

If VolLabel$(1) <> "" Then
    Cmd$ = "label " + FAT32DriveLetter$ + ": " + VolLabel$(1)
End If

' Format partition 2

Shell Cmd$
Cmd$ = "format " + exFATorNTFSdriveletter$ + ": /FS:" + FSType$ + " /Q /Y > NUL"
Shell Cmd$

' Restore the volume label if the original volume label was not blank.

If VolLabel$(2) <> "" Then
    Cmd$ = "label " + exFATorNTFSdriveletter$ + ": " + VolLabel$(2)
    Shell Cmd$
End If

ReDim Letter(2) As String

Letter$(1) = FAT32DriveLetter$
Letter$(2) = exFATorNTFSdriveletter$

GoTo DoneWithBitLocker

' If the user picks the option to wipe the drive and set it up from scratch, then we come here.

AskForPartitions:

AddPart$ = "" ' Set initial value

Cls
Print "We will create two partitions to facilitate making a boot disk that can be booted on both BIOS and UEFI based"
Print "systems. If you want, additional partitions can be created to store other data. Please note that you can add"
Print "a maximum of 2 additional partitions, for a total of 4 partitions."
Print
Input "Do you want to create additional partitions"; AddPart$

' Parse the users response to determine if it is a valid yes / no response.

YesOrNo AddPart$
AddPart$ = YN$

If AddPart$ = "X" Then
    Print
    Color 14, 4: Print "Please provide a valid response.": Color 15
    Pause
    GoTo AskForPartitions
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

' For each additional partition, except the first and last partition, ask for the size.
' The first partition will be 2560 MB and the last partition will be assigned all remaining space.
' Note: 2560 MB may seem excesive for the first partition, but with some customized configurations,
' especially dual architecture, a good amount of space is needed.

Cls
Print "On the next screen you will asked for partition sizes. Note that the first partition will always be created with"
Print "a size of 2.5 GB. You should make the second partition large enough to hold your Windows image. For example, if"
Print "your ISO image is 8 GB in size, make this partition at least 8 GB in size."
Print
Print "TIP: If you plan to update your ISO image in the future, or use a different ISO image, you might want to create"
Print "the second partition with plenty of free space to accomodate any larger images in the future."
Pause

PartitionSizes:

TotalPartitions = AdditionalPartitions + 2

' Get partition sizes. We don't have to ask about the first partition. It will always be 2560 MB.
' We need to remove the leading space for the partition size so we are going to convert it
' to a string. If there are only 2 partitions then we do not need to ask for partition sizes
' since the last partition will be set to occupy all remaining space on the drive. In addition,
' we don't need to ask about encryption.

PartitionSize$(1) = "2560"

If TotalPartitions = 2 Then GoSub AfterBitLockerInfo

For x = 2 To (TotalPartitions - 1)

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

GoSub SelectDisk

' We are done with the variable ListOfDisks$. Let's free up the space it used by clearing it.

ListOfDisks$ = ""

' We are now going to check to see if the selected disk is larger than 2 TB. If it is, then the user needs to either:

' 1) Decide that no space above 2 TB is needed.
' 2) Forfeit the ability to boot on a BIOS based system.

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

' Add up the size of partitions specified by user. The media needs to have more space than what user specified.
' Since partition 1 will always be 2.5 GB we take the 2.5 GB as a starting point and then add the space specified
' for any additional partitions that the user wants to add.

' Init variable

TotalSpaceNeeded = 2560

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

' 2,097,152 = The number of MB in 2 TB.

If (AvailableSpace > 2097152) Then

    AskForOverride:

    Cls
    Print "Do you want to set an MBR override? Type HELP for information about this option."
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

    If Override$ = "X" Then
        Print
        Color 14, 4: Print "Please provide a valid response.": Color 15
        Pause
        GoTo AskForOverride
    End If

End If

Removable:

' Write the commands to initialize the disk to the file named "TEMP.BAT"

Cls
Print "Initializing disk..."
Open "TEMP.BAT" For Output As #1
Print #1, "@echo off"
Print #1, "(echo select disk"; DiskID
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
Print "For each partition, please provide a volume label to assign to the partition. To accept the default name,"
Print "simply press enter. For no volume label, enter the text "; Chr$(34); "NO-LABEL"; Chr$(34); "."
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

    ' If user typed "NO-LABEL" then the new volume label willbe blank.

    If UCase$(NewLabel$) = "NO-LABEL" Then NewLabel$ = ""

    ' If x=1 then we are working on the first volume label which is limited to 11 characters. Anything  after 1 will be either an exFAT partition
    ' which also has an 11 character limit, or NTFS which is limited to 32 characters. We are using the variables "Row" "RowEnd and "Column" to position the cursor on the screen. We do
    ' this because if the user enters an invalid value, we want erase the invalid response from the screen and move the prompt back to
    ' the same place on the screen.

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
    If NewLabel$ <> "" Then
        VolLabel(x) = NewLabel$
    Else
        VolLabel(x) = ""
    End If

    ValidVolLabel:

Next x

Cls
Print
Color 0, 10: Print "NOTE:";: Color 15: Print " If a message pops up saying that a disk needs to be formatted, please click "; Chr$(34); "Cancel"; Chr$(34); "."
Print
Print "Preparing disk. Note that this may take a while."
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
                Print #1, "echo format fs=fat32 quick"
            Else
                Print #1, "echo format fs=fat32 quick label=" + Chr$(34) + VolLabel$(x) + Chr$(34)
            End If
            Print #1, "echo assign letter="; Letter$(x)
        Else
            If VolLabel$(x) = "" Then
                Print #1, "echo format FS="; FSType$; " quick"
            Else
                Print #1, "echo format FS="; FSType$; " quick label=" + Chr$(34) + VolLabel$(x) + Chr$(34)
            End If
            Print #1, "echo assign letter="; Letter$(x)
        End If
    Else
        Print #1, "echo create partition primary"
        If VolLabel$(x) = "" Then
            Print #1, "echo format FS="; FSType$; " quick"
        Else
            Print #1, "echo format FS="; FSType$; " quick label=" + Chr$(34) + VolLabel$(x) + Chr$(34)
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
        Print #1, "echo ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #1, "echo :: We will wait a maximum of 180 seconds for BitLocker to initialize. ::"
        Print #1, "echo :: Typically, BitLocker will initialize much quicker, but we are      ::"
        Print #1, "echo :: allowing time to accomodate very slow media.                       ::"
        Print #1, "echo ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
        Print #1, "echo."
        Print #1, ""
        Print #1, "for /L %%a in (1,1,90) do ("
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
Print #ff, "del WIM_File_Copy_Error.txt > NUL 2>&1"
Print #ff, "echo."
Print #ff, "echo *************************************************************"
Print #ff, "echo * Copying files. Be aware that this can take quite a while, *"
Print #ff, "echo * especially on the 2nd partition and with slower media.    *"
Print #ff, "echo * Please be patient and allow this process to finish.       *"
Print #ff, "echo *************************************************************"
Print #ff, "echo."
Print #ff, "echo Copying files to partition #1"

If ExcludeAutounattend$ = "N" Then
    Print #ff, "robocopy "; CDROM$; "\ "; Letter$(1); ":\ /mir /xd sources /njs /256 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Else
    Print #ff, "robocopy "; CDROM$; "\ "; Letter$(1); ":\ /mir /xf autounattend.xml /xd sources /njs /256 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
End If

Print #ff, "robocopy "; CDROM$; "\sources "; Letter$(1); ":\sources boot.wim /njh /njs /256 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "echo Copying files to partition #2"
Print #ff, "robocopy "; CDROM$; "\sources "; Letter$(2); ":\sources /mir /njh /njs /xf boot.wim /256 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "echo. > "; Letter$(1); ":\PAR1_MEDIA.WIM"
Print #ff, "echo. > "; Letter$(2); ":\PAR2_MEDIA.WIM"
Print #ff, "goto cleanup"
Print #ff, ":HandleError"
Print #ff, "echo An error ocurred > WIM_File_Copy_Error.txt"
Print #ff, ":cleanup"
Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'"; MakeBootableSourceISO$; "'"; Chr$(34); Chr$(34) + " > NUL"
Close #ff
Shell "TEMP.BAT"
If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

' Check for the existance of a file named "WIM_File_Copy_Error.txt". If such a file exists, it indicates that
' that there was an error copying files with the above batch file. In that case, take the following actions:
'
' 1) Display a warning to the user.
' 2) Delete the "WIM_File_Copy_Error.txt" file.
' 3) Abort this routine and return to the start of the program.

If _FileExists("WIM_File_Copy_Error.txt") Then
    Cls
    Print
    Color 14, 4: Print "WARNING!";: Color 15: Print " There was an error copying files. This usually indicates that there was not enough space on the"
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
Print #ff, "del WIM_File_Copy_Error.txt > NUL 2>&1"
Print #ff, "echo."
Print #ff, "echo *************************************************************"
Print #ff, "echo * Copying files. Be aware that this can take quite a while, *"
Print #ff, "echo * especially on the 2nd partition and with slower media.    *"
Print #ff, "echo * Please be patient and allow this process to finish.       *"
Print #ff, "echo *************************************************************"
Print #ff, "echo."
Print #ff, "echo Copying files to partition #1"

If ExcludeAutounattend$ = "N" Then
    Print #ff, "robocopy "; CDROM$; "\ "; Letter$(1); ":\ /mir /xd sources /njs /256 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Else
    Print #ff, "robocopy "; CDROM$; "\ "; Letter$(1); ":\ /mir /xf autounattend.xml /xd sources /njs /256 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
End If

Print #ff, "robocopy "; CDROM$; "\x64\sources "; Letter$(1); ":\x64\sources boot.wim /njh /njs /256 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x86\sources "; Letter$(1); ":\x86\sources boot.wim /njh /njs /256 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "echo Copying files to partition #2"
Print #ff, "robocopy "; CDROM$; "\x64\sources "; Letter$(2); ":\x64\sources /mir /njh /njs /xf boot.wim /256 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "robocopy "; CDROM$; "\x86\sources "; Letter$(2); ":\x86\sources /mir /njh /njs /xf boot.wim /256 > NUL"
Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
Print #ff, "goto cleanup"
Print #ff, ":HandleError"
Print #ff, "echo An error ocurred > WIM_File_Copy_Error.txt"
Print #ff, ":cleanup"
Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'"; MakeBootableSourceISO$; "'"; Chr$(34); Chr$(34)
Close #ff
Shell "TEMP.BAT"
If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

' Check for the existance of a file named "WIM_File_Copy_Error.txt". If such a file exists, it indicates that
' that there was an error copying files with the above batch file. In that case, take the following actions:
'
' 1) Display a warning to the user.
' 2) Delete the "WIM_File_Copy_Error.txt" file.
' 3) Abort this routine and return to the start of the program.

If _FileExists("WIM_File_Copy_Error.txt") Then
    Cls
    Print
    Color 14, 4: Print "WARNING!";: Color 15: Print " There was an error copying files. This usually indicates that there was not enough space on the"
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
Print

' In order to display a more descriptive closing message, we need to know whether the user chose to wipe the entire disk
' or whether they simply chose to refresh the the image on a disk. The variable "WipeOrRefresh" will be set to 1 if they
' chose to wipe the disk, or it will be 2 if they chose to perform a refresh.

Select Case WipeOrRefresh
    Case 1
        Print "The disk that you selected was wiped and two partitions were created for the purpose of making your Windows image"
        Print "bootable."
        Print

        If AddPart$ = "Y" Then
            Print AdditionalPartitions; "additional partition(s) were also created."
        End If

    Case 2
        Print "The Windows image on your disk has been updated using the image that you provided. If the disk contains other"
        Print "partitions, these were not altered in any way."
End Select
Pause

ChDir ProgramStartDir$: GoTo BeginProgram

' The portion of the routine for creating multiple bootable partitions for UEFI based systems begins here.

Boot_MultipleMode:

TotalPartitions = 0
OS_Partitions = 0
PE_Partitions = 0
Other_Partitions = 0

' IMPORTANT: There are two variables you need to have a clear understanding of. Below, I ask for the number of operating systems the user wants to add.
' I put that in OS_Partitions and then double that value because each OS requires 2 partitions. There are places in the code where I do calculations
' that may look a little odd to go back to actual number of operating systems but based upon the number of OS partitions. This got confusing so I've
' added another variable now called OS_Count that simply holds the number of operating systems. We may eventually go back and clean this up.

Cls
Input "How many Operating Systems (Windows 10 or Windows 11) would you like to add"; OS_Count

If OS_Count > 0 Then
    OS_Partitions = OS_Count * 2
Else
    OS_Partitions = 0
End If

TotalPartitions = OS_Partitions

If TotalPartitions > 10 Then
    Cls
    Print "Each operating system creates 2 partitions on the boot media. Currently this program will only accomodate a"
    Print "maximum of 10 total partitions."
    Pause
    GoTo Boot_MultipleMode
End If

Cls
Print "How many Windows PE / RE based programs would you like to add? These are programs such as the Macrium Reflect"
Print "recovery disk or the Acronis TrueImage recovery disk, etc."
Print
Input "How many Windows PE / RE based programs would you like to add"; PE_Partitions
TotalPartitions = TotalPartitions + PE_Partitions

If TotalPartitions > 10 Then
    Cls
    Print "Each operating system creates 2 partitions on the boot media. Each of your Windows PE / RE programs will occupy"
    Print " one additional partition. Currently this program will only accomodate a maximum of 10 partitions."
    Pause
    GoTo Boot_MultipleMode
End If

Cls
Input "How many general purpose partitions would you like to add"; Other_Partitions
TotalPartitions = TotalPartitions + Other_Partitions

If TotalPartitions > 10 Then
    Cls
    Print "Each operating system creates 2 partitions on the boot media. Each of your Windows PE / RE programs will occupy"
    Print " one additional partition as will the general purpose partitions. Currently this program will only accomodate a"
    Print "maximum of 10 partitions."
    Pause
    GoTo Boot_MultipleMode
End If

ReDim _Preserve SourcePath_Multi$(TotalPartitions)
ReDim _Preserve PartitionDescription(TotalPartitions)
ReDim _Preserve ParSize(TotalPartitions)
ReDim _Preserve ParType$(TotalPartitions)
PartitionCounter = 0

If OS_Partitions > 0 Then
    For x = 1 To (OS_Partitions) Step 2
        PartitionCounter = PartitionCounter + 1

        Redo_OS_Partitions:
        Do
            Cls
            Print "Please enter a friendly name for Operating System #"; Int((x + 1) / 2)
            Print
            Input "Friendly Name: ", PartitionDescription$(x)
        Loop While PartitionDescription$(x) = ""

        If Len(PartitionDescription$(x)) > 45 Then
            Print
            Print "Please enter a description of 45 characters or less."
            Pause
            GoTo Redo_OS_Partitions
        End If

        PartitionDescription$(x + 1) = PartitionDescription$(x)

        Redo_OS_Path:

        Cls
        Print "Please enter the full path and file name of the ISO image for ";: Color 0, 10: Print PartitionDescription$(x);: Color 15: Input ": ", SourcePath_Multi$(x)

        If Not _FileExists(SourcePath_Multi$(x)) Then
            Cls
            Print "No such file exists. Please enter a valid path with file name."
            Pause
            GoTo Redo_OS_Path
        End If

        SourcePath_Multi$(x + 1) = SourcePath_Multi$(x)

        Do
            Cls
            Print "Since this is an Operating System Installation Media, we need two partitions to support it. The first partition"
            Print "will be a relativly small partition. I would suggest using a size of 2 GB to start. If you are using a modified"
            Print "Windows image with updates injected, you may need something larger like 2.5 GB."
            Print
            Print "Enter the size of the ";: Color 0, 10: Print "first";: Color 15: Print " partition for ";: Color 0, 10: Print PartitionDescription(x): Color 15
            Print
            GoSub Generic_Partition_Size
            ParSize(x) = Val(ParSizeInMB$)
            ParType$(x) = "FAT32"
        Loop While ParSize(x) = 0

        PartitionCounter = PartitionCounter + 1
        If PartitionCounter = TotalPartitions Then
            Do
                Cls
                Print "This is the second partition for an Operating System boot and is the last partition being created. Since this is"
                Print "the last partition, we can size it to automatically occupy all remaining space on the drive."
                Print
                Input "Do you want to auto size this partition"; AutoSize$
                YesOrNo AutoSize$
                AutoSize$ = YN$
            Loop While AutoSize$ = "X"

            If AutoSize$ = "Y" Then
                ParSize(x + 1) = 0
                ParType$(x + 1) = "NTFS"
                _Continue
            End If
        End If

        Do
            Cls
            Print "This is the second partition for an Operating System."
            Print
            Print "This partition will hold the bulk of the image so it suggested to make it at least as large as your ISO image."
            Print "file. If you think that you may want to manually update this partition with a larger image in the future, then"
            Print "making this partition larger to leave space to grow."
            Print
            Print "Enter the size of the ";: Color 0, 10: Print "second";: Color 15: Print " partition for ";: Color 0, 10: Print PartitionDescription$(x + 1): Color 15
            Print
            GoSub Generic_Partition_Size
            ParSize(x + 1) = Val(ParSizeInMB$)
            ParType$(x + 1) = "NTFS"
        Loop While ParSize(x + 1) = 0

    Next x
End If

' Get PE/RE media details

' If User added any Windows operating systems, then we need to know how many partitions were used by those so that we
' we add WinPE media afterward. We will store that number of partitions as an "Offset".

Offset = 0 ' Set an intial value
If OS_Partitions > 0 Then Offset = OS_Partitions

If PE_Partitions > 0 Then
    For x = 1 To (PE_Partitions)
        PartitionCounter = PartitionCounter + 1

        Redo_PE_Partitions:

        Do
            Cls
            Print "Please enter a friendly name for Windows PE / RE based media number"; x
            Print
            Input "Friendly Name: ", PartitionDescription$(x + Offset)
        Loop While PartitionDescription$(x + Offset) = ""

        If Len(PartitionDescription$(x + Offset)) > 45 Then
            Print
            Print "Please enter a description of 45 characters or less."
            Pause
            GoTo Redo_PE_Partitions
        End If

        Redo_PE_Path:

        Cls
        Print "Please enter the full path and file name of the ISO image for ";: Color 0, 10: Print PartitionDescription$(x + Offset);: Color 15: Input ": ", SourcePath_Multi$(x + Offset)
        If Not _FileExists(SourcePath_Multi$(x + Offset)) Then
            Cls
            Print "No such file exists. Please enter a valid path with file name."
            Pause
            GoTo Redo_PE_Path
        End If

        If PartitionCounter = TotalPartitions Then
            Do
                Cls
                Print "The partition being created for ";: Color 0, 10: Print PartitionDescription(x + Offset);: Color 15: Print " is the last partition in your project."
                Print "Since this is the last partition, we can size it to automatically occupy all remaining space on the drive."
                Print
                Input "Do you want to auto size this partition"; AutoSize$
                YesOrNo AutoSize$
                AutoSize$ = YN$
            Loop While AutoSize$ = "X"

            If AutoSize$ = "Y" Then
                ParSize(x + Offset) = 0
                ParType$(x + Offset) = "FAT32"
                _Continue
            End If
        End If

        Do
            Cls
            Print "Enter the size of the partition for ";: Color 0, 10: Print PartitionDescription(x + Offset): Color 15
            Print
            GoSub Generic_Partition_Size
            ParSize(x + Offset) = Val(ParSizeInMB$)
            ParType$(x + Offset) = "FAT32"
        Loop While ParSize(x + Offset) = 0

    Next x
End If

' Update the partition offset. The whole idea of this process is to make bootable media, so there should be either previous Operating System or WinPE
' based partitions or both specified already. However, we'll verify that now and create the proper offset.

Offset = Offset + PE_Partitions
If Offset = 0 Then
    Cls
    Print "You created no Operating system partitions and no WinPE based media partitions. This program can still create other"
    Print "partitions for you and BitLocker encrypt them for you."
    Pause
End If

If Other_Partitions > 0 Then
    For x = 1 To (Other_Partitions)
        PartitionCounter = PartitionCounter + 1

        Redo_Other_Partitions:

        Do
            Cls
            Print "Please enter a friendly name for the general purpose partition number"; x
            Print
            Input "Friendly Name: ", PartitionDescription$(x + Offset)
        Loop While PartitionDescription$(x + Offset) = ""

        If Len(PartitionDescription$(x + Offset)) > 45 Then
            Print
            Print "Please enter a description of 45 characters or less."
            Pause
            GoTo Redo_Other_Partitions
        End If

        If PartitionCounter = TotalPartitions Then
            Do
                Cls
                Print "The partition being created for ";: Color 0, 10: Print PartitionDescription(x + Offset);: Color 15: Print " is the last partition in your project."
                Print "Since this is the last partition, we can size it to automatically occupy all remaining space on the drive."
                Print
                Input "Do you want to auto size this partition"; AutoSize$
                YesOrNo AutoSize$
                AutoSize$ = YN$
            Loop While AutoSize$ = "X"

            If AutoSize$ = "Y" Then
                ParSize(x + Offset) = 0
                ParType$(x + Offset) = "NTFS"
                _Continue
            End If
        End If

        Do
            Cls
            Print "Enter the size of the partition for ";: Color 0, 10: Print PartitionDescription(x + Offset): Color 15
            Print
            GoSub Generic_Partition_Size
            ParSize(x + Offset) = Val(ParSizeInMB$)
            ParType$(x + Offset) = "NTFS"
        Loop While ParSize(x + Offset) = 0
    Next x
End If

If TotalPartitions = 0 Then
    Cls
    Print "You specified zero partitions. We will now return you to the main menu so that you can ponder what you have done."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

Do
    Cls
    Print "Here is a summary of what we have so far:"
    Print
    Print "                  ";: Color 0, 10: Print "Description";: Color 15: Print "                                 ";: Color 0, 10: Print "Partition Size (in MB)";
    Color 15: Print "   ";: Color 0, 10: Print "Partition Type": Color 15
    Print

    If OS_Partitions > 0 Then
        For x = 1 To (Int(OS_Partitions / 2))
            Print PartitionDescription$((x * 2) - 1); " boot partition";: Locate CsrLin, 70: Print ParSize((x * 2) - 1);: Locate CsrLin, 92: Print "FAT32"
            Print PartitionDescription$(x * 2); " setup partition";: Locate CsrLin, 70: Print ParSize(x * 2);: Locate CsrLin, 92: Print "NTFS"
        Next x
    End If

    If PE_Partitions > 0 Then
        For x = 1 To PE_Partitions
            Print PartitionDescription$(OS_Partitions + x); " boot partition";: Locate CsrLin, 70: Print ParSize(OS_Partitions + x);: Locate CsrLin, 92: Print "FAT32"
        Next x
    End If

    If Other_Partitions > 0 Then
        For x = 1 To Other_Partitions
            Print PartitionDescription$(OS_Partitions + PE_Partitions + x); " NTFS partition";: Locate CsrLin, 70
            Print ParSize(OS_Partitions + PE_Partitions + x);: Locate CsrLin, 92: Print "NTFS"
        Next x
    End If

    Print
    Print "If the last partition shows a size of 0, this indicates that the partition is set to occupy all remaining space."
    Print
    Input "Is the above information correct"; Temp$
    YesOrNo Temp$
Loop While YN$ = "X"
Temp$ = YN$

If Temp$ = "N" Then
    Cls
    Print "Please organize your information and run this routine again."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

' Present a list of disks to the user so that they can select the disk to be used for this project.

GoSub SelectDisk

' Verify that the disk selected has enough space to hold all the partitions requested by the user.

' Verify that the amount of space available is greater than what is specified by the user

Cls
Print "Perform initial preparation of the selected drive..."
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

TotalSpaceNeeded = 0

For x = 1 To TotalPartitions
    TotalSpaceNeeded = TotalSpaceNeeded + ParSize(x)

    ' For the last partition, a size of zero indicates that all remaining space should be used. In the event that we encounter
    ' a size of zero, we should set aside a minimum of 100MB since that is the minimum partition size we are enforcing

    If ParSize(x) = 0 Then
        TotalSpaceNeeded = TotalSpaceNeeded + 100
    End If
Next x

If TotalSpaceNeeded > AvailableSpace Then
    Cls
    Color 14, 4: Print "Warning!";: Color 15: Print " You have have specified partition sizes that total more than the space available on the selected disk."
    Print
    Print "Please check the values that you have supplied and the disk that you selected and try again."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

' Perform a simple check to validate that all OS images specified appear to be valid Win x64 images.

If OS_Partitions > 0 Then
    For x = 1 To OS_Partitions
        DetermineArchitecture SourcePath_Multi$(x), 1
        Select Case ImageArchitecture$
            Case "x86", "DUAL", "NONE"
                Cls
                Print
                Print "The file specified for "; PartitionDescription$(x); " appears to be invalid."
                Print "The image specified must be a Windows x64 image. This particular routine does not support x86 or dual"
                Print "architecture images, nor does it support other types of images."
                Print
                Print "Please resolve this issue and then run this program again."
                Pause
                ChDir ProgramStartDir$: GoTo BeginProgram
            Case "x64"
                ' Do nothing - We want the image to be of type x64 so we'll proceed on at this point.
        End Select
    Next x
End If

' Currently, we are not performing any validation on Win PE / RE media. This may be considered for future addition.

' We will now ask the user if they want to manually or automatically select drive letter assignments for the partitions being created.
' If auto, we'll get those drive letters now.

ff = FreeFile
Open "temp.bat" For Output As #ff
Print #ff, "@echo off"
Print #ff, "(echo select disk"; DiskID
Print #ff, "echo clean"
Print #ff, "echo convert gpt"
Print #ff, "echo exit"
Print #ff, ") | diskpart > NUL"
Close #ff
Shell _Hide "temp.bat"
If _FileExists("temp.bat") Then Kill "temp.bat"

GoSub SelectAutoOrManual

Do
    Cls
    Print "Here is an updated summary along with drive letter assignments:"
    Print
    Print "                  ";: Color 0, 10: Print "Description";: Color 15: Print "                                 ";: Color 0, 10: Print "Partition Size (in MB)";: Color 15
    Print "   ";: Color 0, 10: Print "Partition Type";: Color 15: Print "   ";: Color 0, 10: Print "Drive Letter": Color 15
    Print

    If OS_Partitions > 0 Then
        For x = 1 To (Int(OS_Partitions / 2))
            Print PartitionDescription$((x * 2) - 1); " boot partition";: Locate CsrLin, 70: Print ParSize((x * 2) - 1);
            Locate CsrLin, 92: Print "FAT32";: Locate CsrLin, 110: Print Letter$((x * 2) - 1); ":"
            Print PartitionDescription$(x * 2); " setup partition";: Locate CsrLin, 70: Print ParSize(x * 2);
            Locate CsrLin, 92: Print "NTFS";: Locate CsrLin, 110: Print Letter$(x * 2); ":"
        Next x
    End If

    If PE_Partitions > 0 Then
        For x = 1 To PE_Partitions
            Print PartitionDescription$(OS_Partitions + x); " boot partition";: Locate CsrLin, 70: Print ParSize(OS_Partitions + x);
            Locate CsrLin, 92: Print "FAT32";: Locate CsrLin, 110: Print Letter$(OS_Partitions + x); ":"
        Next x
    End If

    If Other_Partitions > 0 Then
        For x = 1 To Other_Partitions
            Print PartitionDescription$(OS_Partitions + PE_Partitions + x); " NTFS partition";: Locate CsrLin, 70: Print ParSize(OS_Partitions + PE_Partitions + x);
            Locate CsrLin, 92: Print "NTFS";: Locate CsrLin, 110: Print Letter$(OS_Partitions + PE_Partitions + x); ":"
        Next x
    End If

    Print
    Print "If the last partition shows a size of 0, this indicates that the partition is set to occupy all remaining space."
    Print
    Input "Is the above information correct"; Temp$
    YesOrNo Temp$
Loop While YN$ = "X"
Temp$ = YN$

If Temp$ = "N" Then
    Cls
    Print "Please organize your information and run this routine again."
    Pause
    ChDir ProgramStartDir$: GoTo BeginProgram
End If

' Ask user if they want to hide drive letters for OS and Win PE / RE partitions.

HideLetters$ = "N" ' Set an initial value

If (OS_Partitions + PE_Partitions) > 0 Then
    Do
        Cls
        Print "You have chosen to create bootable partitions. Normally, you would only need to interact with these when booting"
        Print "your system from them. If you wish, the program can remove the drive letters from these partitions to keep things"
        Print "tidy and free up some drive letters. The drive will still be bootable."
        Print
        Print "Note that the drive letters will only be hidden on this system. If you take the drive to another system it may"
        Print "assign letters. Also, you can always unhide these partitions by simply opening Disk Manager and assigning letters."
        Print
        Input "Do you want to hide drive letters for operating system and Win PE / RE partitions"; HideLetters$
        YesOrNo HideLetters$
        HideLetters$ = YN$
    Loop While HideLetters$ = "X"
End If

Cls
Print "Preparing the selected disk now. Please standby..."
Print
ff = FreeFile
Open "temp.bat" For Output As #ff
Print #ff, "@echo off"
Print #ff, "(echo select disk"; DiskID

For x = 1 To TotalPartitions

    ' If the specified partition size is 0, this indicates that we want to allow that partition to occupy all remaining space.
    ' Handle the sizing of this partition accordingly. Leving off the size= parameter in diskpart will cause all remaining
    ' space to be used.

    If ParSize(x) = 0 Then
        Print #ff, "echo create partition primary"
    Else
        Print #ff, "echo create partition primary size="; LTrim$(Str$(ParSize(x)))
    End If

    Print #ff, "echo format fs="; ParType$(x); " quick label=PAR"; LTrim$(Str$(x))
    Print #ff, "echo assign letter="; Letter$(x)
Next x

Print #ff, "echo exit"
Print #ff, ") | diskpart > NUL"
Close #ff

Shell _Hide "temp.bat"

If _FileExists("temp.bat") Then Kill "temp.bat"

' Process OS partitions

If OS_Count = 0 Then GoTo No_OS_Partitions

PartitionCounter = 1
For x = 1 To OS_Count
    Cls
    Print "Working on OS image"; x; "of"; OS_Count
    Color 0, 10: Print PartitionDescription(PartitionCounter): Color 15
    Print

    ' mount the windows image.

    MountISO SourcePath_Multi$(PartitionCounter)
    CDROM$ = MountedImageDriveLetter$

    ' Copy all files except the \sources folder to the first partition
    ' make a directory called sources and copy only the boot.wim
    ' on the first partition, create a folder called \sources. Copy all files from the original \sources EXCEPT boot.wim
    ' on the second partition, create a folder called \sources. Copy all files from the original \sources EXCEPT boot.wim

    ff = FreeFile
    Open "TEMP.BAT" For Output As #ff
    Print #ff, "@echo off"
    Print #ff, "del WIM_File_Copy_Error.txt > NUL 2>&1"
    Print #ff, "echo."
    Print #ff, "echo *************************************************************"
    Print #ff, "echo * Copying files. Be aware that this can take quite a while, *"
    Print #ff, "echo * especially on the 2nd partition and with slower media.    *"
    Print #ff, "echo * Please be patient and allow this process to finish.       *"
    Print #ff, "echo *************************************************************"
    Print #ff, "echo."
    Print #ff, "echo Copying files to partition #"; LTrim$(Str$(PartitionCounter))

    If ExcludeAutounattend$ = "N" Then
        Print #ff, "robocopy "; CDROM$; "\ "; Letter$(PartitionCounter); ":\ /mir /xd sources /njs /256 > NUL"
        Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
    Else
        Print #ff, "robocopy "; CDROM$; "\ "; Letter$(PartitionCounter); ":\ /mir /xf autounattend.xml /xd sources /njs /256 > NUL"
        Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
    End If

    Print #ff, "robocopy "; CDROM$; "\sources "; Letter$(PartitionCounter); ":\sources boot.wim /njh /njs /256 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
    Print #ff, "echo Copying files to partition #"; LTrim$(Str$(PartitionCounter + 1))
    Print #ff, "robocopy "; CDROM$; "\sources "; Letter$(PartitionCounter + 1); ":\sources /mir /njh /njs /xf boot.wim /256 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
    Print #ff, "goto cleanup"
    Print #ff, ":HandleError"
    Print #ff, "echo An error ocurred > WIM_File_Copy_Error.txt"
    Print #ff, ":cleanup"
    Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'"; SourcePath_Multi$(PartitionCounter); "'"; Chr$(34); Chr$(34) + " > NUL"
    Close #ff
    Shell "TEMP.BAT"
    If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

    ' Check for the existance of a file named "WIM_File_Copy_Error.txt". If such a file exists, it indicates that
    ' that there was an error copying files with the above batch file. In that case, take the following actions:
    '
    ' 1) Display a warning to the user.
    ' 2) Delete the "WIM_File_Copy_Error.txt" file.
    ' 3) Abort this routine and return to the start of the program.

    If _FileExists("WIM_File_Copy_Error.txt") Then
        Cls
        Print
        Color 14, 4: Print "WARNING!";: Color 15: Print " There was an error copying files. This usually indicates that there was not enough space on the"
        Print "destination. Please correct this situation and run this routine again."
        Pause
        ChDir ProgramStartDir$: GoTo BeginProgram
    End If

    ' Making the file ei.cfg on the partition 2, in the sources folder if the user wanted this file added.
    ' If an ei.cfg already exists, leave it alone.

    If CreateEiCfg$ = "Y" Then
        Temp$ = Letter$(PartitionCounter + 1) + ":\sources\ei.cfg"

        If Not (_FileExists(Temp$)) Then
            ff = FreeFile
            Open (Temp$) For Output As #ff
            Print #ff, "[CHANNEL]"
            Print #ff, "Retail"
            Close #ff
        End If

    End If

    PartitionCounter = PartitionCounter + 2

Next x

No_OS_Partitions:

' Process any Win PE/RE Images

PartitionCounter = OS_Partitions + 1

If PE_Partitions = 0 Then GoTo No_PE_Partitions

For x = 1 To PE_Partitions
    Cls
    Print "Working on PE/RE image"; x; "of"; PE_Partitions
    Color 0, 10: Print PartitionDescription(PartitionCounter): Color 15
    Print

    ' mount the windows image.

    MountISO SourcePath_Multi$(PartitionCounter)
    CDROM$ = MountedImageDriveLetter$

    ' Copy ALL files from source to destination

    ff = FreeFile
    Open "TEMP.BAT" For Output As #ff
    Print #ff, "@echo off"
    Print #ff, "del WIM_File_Copy_Error.txt > NUL 2>&1"
    Print #ff, "echo."
    Print #ff, "echo ************************************************"
    Print #ff, "echo * Copying files. This may take a little while. *"
    Print #ff, "echo ************************************************"
    Print #ff, "echo."
    Print #ff, "echo Copying files to partition #"; LTrim$(Str$(PartitionCounter))
    Print #ff, "robocopy "; CDROM$; "\ "; Letter$(PartitionCounter); ":\ /mir /njs /256 > NUL"
    Print #ff, "if %ERRORLEVEL% gtr 3 goto HandleError"
    Print #ff, "goto cleanup"
    Print #ff, ":HandleError"
    Print #ff, "echo An error ocurred > WIM_File_Copy_Error.txt"
    Print #ff, ":cleanup"
    Print #ff, "powershell.exe -command "; Chr$(34); "Dismount-DiskImage "; Chr$(34); "'"; SourcePath_Multi$(PartitionCounter); "'"; Chr$(34); Chr$(34) + " > NUL"
    Close #ff

    Shell "TEMP.BAT"
    If _FileExists("TEMP.BAT") Then Kill "TEMP.BAT"

    ' Check for the existance of a file named "WIM_File_Copy_Error.txt". If such a file exists, it indicates that
    ' that there was an error copying files with the above batch file. In that case, take the following actions:
    '
    ' 1) Display a warning to the user.
    ' 2) Delete the "WIM_File_Copy_Error.txt" file.
    ' 3) Abort this routine and return to the start of the program.

    If _FileExists("WIM_File_Copy_Error.txt") Then
        Cls
        Print
        Color 14, 4: Print "WARNING!";: Color 15: Print " There was an error copying files. This usually indicates that there was not enough space on the"
        Print "destination. Please correct this situation and run this routine again."
        Pause
        ChDir ProgramStartDir$: GoTo BeginProgram
    End If

    PartitionCounter = PartitionCounter + 1

Next x

No_PE_Partitions:

' No actions are needed for the generic partitions as they were already created previously
' and no data needs to be copied to these partitions.

If HideLetters$ = "N" Then GoTo DoneHidingLetters

For x = 1 To (OS_Partitions + PE_Partitions)
    Shell "mountvol " + Letter$(x) + ": /D"
Next x

DoneHidingLetters:

' At this time, all operations are completed and we return to the main menu.

Cls
Print "Operations completed. Please be aware that when you boot from this media, the UEFI menu will not display friendly names"
Print "for each boot entry. The names shown will be based upon the hardware device. The sample below shows an example of what"
Print "this might look like. It will vary from system to system. Note that in this example, the three lines that show "; Chr$(34); "USB"
Print "Hard Drive(UEFI) - SanDisk Extreme Pro 0"; Chr$(34); " are our bootable partitions."
Print
Print "    OS Boot Manager (UEFI) - Windows Boot Manager (Seagate Firecuda 510 SSD ZP1000GM30001)"
Print "    USB Hard Drive(UEFI) - SanDisk Extreme Pro 0"
Print "    USB Hard Drive(UEFI) - SanDisk Extreme Pro 0"
Print "    USB Hard Drive(UEFI) - SanDisk Extreme Pro 0"
Print "    Boot From EFI File"
Print
Print "Note that on some systems you will see one item for each bootable FAT32 partition created on your disk, while some"
Print "system will display an entry for each FAT32 <AND> NTFS partition. You will need to keep track of these and select the"
Print "FAT32 partitions to boot from."
Print
Pause
Cls
Print "Here is a summary of how we configured your disk:"
Print
Print "                  ";: Color 0, 10: Print "Description";: Color 15: Print "                                 ";: Color 0, 10: Print "Partition Size (in MB)";: Color 15
Print "   ";: Color 0, 10: Print "Partition Type";: Color 15: Print "   ";: Color 0, 10: Print "Drive Letter": Color 15
Print

If OS_Partitions > 0 Then
    For x = 1 To (Int(OS_Partitions / 2))
        Print PartitionDescription$((x * 2) - 1); " boot partition";
        Locate CsrLin, 70: Print ParSize((x * 2) - 1);
        Locate CsrLin, 92: Print "FAT32";

        If HideLetters$ = "Y" Then
            Locate CsrLin, 109: Print "NONE"
        Else
            Locate CsrLin, 110: Print Letter$((x * 2) - 1); ":"
        End If

        Print PartitionDescription$(x * 2); " setup partition";
        Locate CsrLin, 70: Print ParSize(x * 2);
        Locate CsrLin, 92: Print "NTFS";

        If HideLetters$ = "Y" Then
            Locate CsrLin, 109: Print "NONE"
        Else
            Locate CsrLin, 110: Print Letter$(x * 2); ":"
        End If

    Next x
End If

If PE_Partitions > 0 Then
    For x = 1 To PE_Partitions
        Print PartitionDescription$(OS_Partitions + x); " boot partition";
        Locate CsrLin, 70: Print ParSize(OS_Partitions + x);
        Locate CsrLin, 92: Print "FAT32";

        If HideLetters$ = "Y" Then
            Locate CsrLin, 109: Print "NONE"
        Else
            Locate CsrLin, 110: Print Letter$(OS_Partitions + x); ":"
        End If

    Next x
End If

If Other_Partitions > 0 Then
    For x = 1 To Other_Partitions
        Print PartitionDescription$(OS_Partitions + PE_Partitions + x); " NTFS partition";
        Locate CsrLin, 70: Print ParSize(OS_Partitions + PE_Partitions + x);
        Locate CsrLin, 92: Print "NTFS";: Locate CsrLin, 110: Print Letter$(OS_Partitions + PE_Partitions + x); ":"
    Next x
End If

Print
Print "If the last partition shows a size of 0, this indicates that the partition is set to occupy all remaining space."
Print
Print "Note that only those partitions described as a boot partition will be displayed on the UEFI boot menu."
Pause

ChDir ProgramStartDir$: GoTo BeginProgram

' End of mainroutine

' Subroutine - Shows patition information.

ShowPartitionSizes:

Cls
Color 0, 14
Print "*******************"
Print "* PARTITION SIZES *"
Print "*******************"
Color 15
Print
Print "Partition #1:";: Color 0, 10: Print "   2.50 GB ";: Color 15: Print "(Holds boot files)"
Print "Partition #2:";
If PartitionSize$(2) = "" Then
    Color 14, 4: Print " NOT YET DEFINED ";: Color 15
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

Print "(Must be large enough to hold contents of Windows image)"

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
        Print "Partition #4:";: Color 0, 10: Print " All remaining space not assigned to the first three partitions ": Color 15
    Case 3
        Print "Partition #3:";: Color 0, 10: Print " All remaining space not assigned to the first two partitions ": Color 15
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
Print "You have selected the following disk: ";: Color 0, 14: Print "Disk"; DiskID; "- "; DiskDetail$(x): Color 15
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
' but we resume checking for free drive letters with the letter after the last one we just assigned. To determine if a
' drive letter is available, we first check to if the drive letter exists using a _DIREXISTS command. The problem with
' this is that _DIREXISTS will indicate that a BitLocker encrypted drive does NOT exist if it is locked. As a result,
' if _DIREXISTS indicates that a drive letter does NOT exist, we follow up by running the command "manage-bde -status D:",
' where D: is the drive letter we want to test. We grab the output of that command and check for the string "could not be
' opened by BitLocker". If we find that string, then that drive letter is not BitLocker encrypted and so it really is a free
' drive letter in that case.

LettersAssigned = 0 ' Keep track of how many drive letters were assigned. Once equal to the number of partitions, we are done.

Restore DriveLetterData

For y = 1 To 24
    Read Letter$(LettersAssigned + 1)

    If Not (_DirExists(Letter$(LettersAssigned + 1) + ":")) Then
        Cmd$ = "manage-bde -status " + Letter$(LettersAssigned + 1) + ": > BitLockerStatus.txt"
        Shell Cmd$
        ff = FreeFile
        Open "BitLockerStatus.txt" For Input As #ff
        FileLength = LOF(ff)
        Temp$ = Input$(FileLength, ff)
        Close #ff
        Kill "BitLockerStatus.txt"

        If InStr(Temp$, "could not be opened by BitLocker") Then
            LettersAssigned = LettersAssigned + 1
        End If

        If LettersAssigned = TotalPartitions Then GoTo LetterAssignmentDone

    End If

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

' ********************************************************************************************
' * Create a bootable Windows ISO image that can include multiple editions and architectures *
' ********************************************************************************************

' NOTE: This code was originally a copy / edit of the code for the section "Inject Windows updates into one or more Windows ISO images"
' with the code for injecting updates removed. As such, there may be some variables or elements left over that don't seem to make sense.
' There may also be some variable names that elude to "updates" even though we are not actually injecting any updates.

MakeMultiBootImage:

' This routine will extract Windows editions from one or more ISO images and combine them into a single multi boot ISO image.

' Ask for source folder. Check to make sure folder contains ISO images. If it does, ask if all ISO images should be processed.
' For each image to be processed, we need to keep track of the image name to be processed. Likewise, we need to track source folder.
' For each file, determine if it is dual architecture, x86, or x64.

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
    Print "Enter the path to one or more Windows ISO image files. These can be x64, x86, or dual architecture. These images must"
    Print "include install.wim file(s), ";: Color 0, 10: Print "NOT";: Color 15: Print " install.esd. ";: Color 0, 10: Print "DO NOT";: Color 15: Print " include a file name or extension."
    Print
    Input "Enter the path: ", SourceFolder$
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
        Print "Please standby for a moment. Verifying the architecture of the following image:"
        Print
        Color 10
        Print FileArray$(TotalFiles)
        Color 15
        Print
        Temp$ = FolderArray$(TotalFiles) + FileArray$(TotalFiles)
        DetermineArchitecture Temp$, 1
        Select Case ImageArchitecture$
            Case "x64", "x86"
                FileSourceType$(TotalFiles) = ImageArchitecture$
            Case "DUAL"
                FileSourceType$(TotalFiles) = "x64_DUAL"
                TotalFiles = TotalFiles + 1

                ReDim _Preserve UpdateFlag(TotalFiles) As String
                ReDim _Preserve FileArray(TotalFiles) As String
                ReDim _Preserve FolderArray(TotalFiles) As String
                ReDim _Preserve FileSourceType(TotalFiles) As String

                UpdateFlag$(TotalFiles) = "Y"
                FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
                FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
                FileSourceType$(TotalFiles) = "x86_DUAL"
            Case "NONE"
                Cls
                Color 14, 4: Print "WARNING!";: Color 15: Print " An invalid file has been selected."
                Print "Check the following file to make sure that it is valid. It needs to contain INSTALL.WIM file(s), not INSTALL.ESD."
                Print
                Print "Path: ";: Color 10: Print Left$(Temp$, ((_InStrRev(Temp$, "\"))) - 1): Color 15
                Print "File: ";: Color 10: Print Right$(Temp$, (Len(Temp$) - (_InStrRev(Temp$, "\")))): Color 15
                Pause
                ChDir ProgramStartDir$: GoTo BeginProgram
        End Select
    Next x
    GoTo MMBICheckForMoreFolders
End If

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

            ' Init variables

            ReDim _Preserve UpdateFlag(TotalFiles) As String
            ReDim _Preserve FileArray(TotalFiles) As String
            ReDim _Preserve FolderArray(TotalFiles) As String
            ReDim _Preserve FileSourceType(TotalFiles) As String
            UpdateFlag$(TotalFiles) = "Y"

            FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
            FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
            Cls
            Print "Please standby for a moment. Verifying the architecture of the following image:"
            Print
            Color 10
            Print FileArray$(TotalFiles)
            Color 15
            Print
            Temp$ = FolderArray$(TotalFiles) + FileArray$(TotalFiles)
            DetermineArchitecture Temp$, 1
            Select Case ImageArchitecture$
                Case "x64", "x86"
                    FileSourceType$(TotalFiles) = ImageArchitecture$
                Case "DUAL"
                    FileSourceType$(TotalFiles) = "x64_DUAL"
                    TotalFiles = TotalFiles + 1

                    ' Init variables

                    ReDim _Preserve UpdateFlag(TotalFiles) As String
                    ReDim _Preserve FileArray(TotalFiles) As String
                    ReDim _Preserve FolderArray(TotalFiles) As String
                    ReDim _Preserve FileSourceType(TotalFiles) As String
                    UpdateFlag$(TotalFiles) = "Y"

                    FileArray$(TotalFiles) = Mid$(TempArray$(x), _InStrRev(TempArray$(x), "\") + 1)
                    FolderArray$(TotalFiles) = Left$(TempArray$(x), (_InStrRev(TempArray$(x), "\")))
                    FileSourceType$(TotalFiles) = "x86_DUAL"
                Case "NONE"
                    Cls
                    Color 14, 4: Print "WARNING!";: Color 15: Print " An invalid file has been selected."
                    Print "Check the following file to make sure that it is valid. It needs to contain INSTALL.WIM file(s), not INSTALL.ESD."
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

' For any dual architecture images, we need to determine if the user wants to add Windows editions in both the x64 and x86 folders.

For x = 1 To TotalFiles
    If (FileSourceType$(x) = "x86_DUAL") Or (FileSourceType$(x) = "x64_DUAL") Then
        Cls
        Print "The file listed below is a dual architecture image. We need to know if you intend to add both x64 and x86"
        Print "editions of Windows or not."
        Print
        Color 10: Print "Filename: "; FileArray$(x): Color 15
        Print
        Print "Do you want to add ";: Color 0, 10: Print "ANY";: Color 15: Print " of the ";: Color 0, 14: Print Left$(FileSourceType$(x), 3);: Color 15: Print " editions within this file";: Input Temp$
        YesOrNo Temp$
        Select Case YN$
            Case "X"
                Cls
                Print
                Color 14, 4
                Print "Please provide a valid response."
                Color 15
                Pause
            Case "N"
                UpdateFlag$(x) = "N"
        End Select
    End If
Next x

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
    If (FileSourceType$(IndexCountLoop) = "x64_DUAL") Or (FileSourceType$(IndexCountLoop) = "x86_DUAL") Then
        Print "*******************************************************"
        Print "* This file is a dual architecture file. Please enter *"
        Print "* the index numbers for the ";: Color 0, 14: Print ">> "; Left$(FileSourceType$(IndexCountLoop), 3); " <<";: Color 15: Print " editions below. *"
        Print "*******************************************************"
        Print
    End If
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

    ' We have to specify the values for the ArchitectureChoice$ if our source is a dual architecture file so that
    ' the path to the install.wim or install.esd is correct.

    Select Case FileSourceType$(IndexCountLoop)
        Case "x64_DUAL"
            ArchitectureChoice$ = "x64"
        Case "x86_DUAL"
            ArchitectureChoice$ = "x86"
        Case Else
            ArchitectureChoice$ = ""
    End Select
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
        Select Case ArchitectureChoice$
            Case ""
                Temp$ = _Trim$(Str$(NumberOfSingleIndices))
                IndexRange$ = "1-" + Temp$
            Case "x64"
                Temp$ = _Trim$(Str$(NumberOfx64Indices))
                IndexRange$ = "1-" + Temp$
            Case "x86"
                Temp$ = _Trim$(Str$(NumberOfx86Indices))
                IndexRange$ = "1-" + Temp$
        End Select
        If IndexRange$ = "1-1" Then IndexRange$ = "1"
    End If
    Kill "WIM_Info.txt"

    MMBIProcessRange:

    ProcessRangeOfNums IndexRange$, 1
    If ValidRange = 0 Then
        Color 14, 4
        Print "You did not enter a valid range of numbers"
        Color 15
        Pause
        GoTo MMBIGetMyIndexList
    End If

    ' We will now get WIM info and save it to a file called WIM_Info.txt. We will parse that file to verify that the index
    ' selected is valid. If not, we will ask the user to choose a valid index.

    SourcePath$ = FolderArray$(IndexCountLoop) + FileArray$(IndexCountLoop)
    Print
    Print "Verifying indices."
    Print
    Print "Please standby..."
    Print
    GetWimInfo_Main SourcePath$, 1

    ' If we are processing a file from a dual architecture image, then we need to make sure that we are only processing the
    ' part of the file that pertains to the x64 or the x86 portion of the file that we need.

    For x = 1 To TotalNumsInArray
        WimInfoFound = 0 ' Init Variable
        DualArchitectureFlag$ = ""
        Open "WIM_Info.txt" For Input As #1
        Do
            Line Input #1, WimInfo$
            If (InStr(WimInfo$, "x86 Editions")) Then DualArchitectureFlag$ = "x86_DUAL"
            If (InStr(WimInfo$, "x64 Editions")) Then DualArchitectureFlag$ = "x64_DUAL"
            If (FileSourceType$(IndexCountLoop) = "x86_DUAL") Or (FileSourceType$(IndexCountLoop) = "x64_DUAL") Then
                If FileSourceType$(IndexCountLoop) <> DualArchitectureFlag$ Then GoTo MMBISkipToNextLine_Section1
            End If
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
    Kill "WIM_Info.txt"

    MMBINoIndex:

Next IndexCountLoop

' Now that we have a valid source directory and we know that there are ISO images located there, ask the user for the location where we should save the project

DestinationFolder$ = "" ' Set initial value

MMBIGetDestinationPath10:

Do
    Cls
    Print "Enter the path where the project will be created. We will use this location to save temporary files and we will also"
    Print "save the final ISO image file here."
    Print
    Input "Enter the path where the project should be created: ", DestinationFolder$
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

' If we have arrived here it means that the destination path already exists or we were able to create it successfully.

' Ask user what they want to name the final ISO image file

Cls
UserSelectedImageName$ = "" ' Set initial value
Print "If you would like to specify a name for the final ISO image file that this project will create, please do so now,"
Print "WITHOUT an extension. You can also simply press ENTER to use the default name of Windows.ISO."
Print
Print "Enter name ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension, or press ENTER: ";: Input "", UserSelectedImageName$

If UserSelectedImageName = "" Then
    UserSelectedImageName$ = "Windows.ISO"
Else
    UserSelectedImageName$ = UserSelectedImageName$ + ".ISO"
End If

' Get a count of x64, x86, and DUAL architecture images. If we have both x64 and x86 images but no
' dual architecture images, then we will need to dynamically create the "bcd" files needed for our
' project. If a dual architecture file is a part of the project, then we can simply use it as it
' contains everything that we need. If we find a dual architecture image, then save the path for
' this image to the variable DualArcImagePath$ so that we can use it for building the base image.

' IMPORTANT: The count of files listed immediately below is the number of files of each type in the folders specified
' INCLUDING FILES THAT WILL NOT BE ADDED TO THE MULTI BOOT IMAGE.

x86FileCount = 0
x64FileCount = 0
DUALFileCount = 0
DualArcImagePath$ = ""

' The next set of variables will hold the actual number of each image type to be processed

x86UpdateImageCount = 0
x64UpdateImageCount = 0

For x = 1 To TotalFiles
    Select Case FileSourceType$(x)
        Case "x64"
            x64FileCount = x64FileCount + 1
            If UpdateFlag$(x) = "Y" Then x64UpdateImageCount = x64UpdateImageCount + IndexCount(x)
        Case "x86"
            x86FileCount = x86FileCount + 1
            If UpdateFlag$(x) = "Y" Then x86UpdateImageCount = x86UpdateImageCount + IndexCount(x)
        Case "x86_DUAL", "x64_DUAL"
            DUALFileCount = DUALFileCount + 1
            If (UpdateFlag$(x) = "Y") And (FileSourceType(x) = "x64_DUAL") Then x64UpdateImageCount = x64UpdateImageCount + IndexCount(x)
            If (UpdateFlag$(x) = "Y") And (FileSourceType(x) = "x86_DUAL") Then x86UpdateImageCount = x86UpdateImageCount + IndexCount(x)
            If DUALFileCount = 1 Then DualArcImagePath$ = FolderArray$(x) + FileArray$(x)
    End Select
Next x

' NOTE: When updating the image count, dual architecture images will count as 2 images since we
' list the x64 and x86 images seperately. As a result, when done getting the count, we will need to
' divide the count for dual architecture images by 2.

DUALFileCount = DUALFileCount / 2
TotalImagesToUpdate = x64UpdateImageCount + x86UpdateImageCount

' Create a flag to indicate if this project will be a single architecture project or dual architecture.

If ((x64UpdateImageCount > 0) And (x86UpdateImageCount > 0)) Then
    ProjectType$ = "DUAL"
Else
    ProjectType$ = "SINGLE"
End If

If ProjectType$ = "SINGLE" Then GoTo MMBIEND_GetDualArcImagePath

' If DualArcImagePath is blank, this means that we do not have a dual architecture image available.

If DualArcImagePath$ <> "" Then
    DualBootPackage = 0
    GoTo MMBIEND_GetDualArcImagePath
Else
    DualBootPackage = 1
    GoTo MMBIEND_GetDualArcImagePath
End If

' If we reach this point, then the image specified by the user is valid.

MMBIEND_GetDualArcImagePath:

' Before starting the update process, verify that there are no leftover files sitting in the destination.

Cleanup DestinationFolder$
If CleanupSuccess = 0 Then ChDir ProgramStartDir$: GoTo BeginProgram

' Create the folders we need for the project.

MkDir DestinationFolder$ + "ISO_Files"
MkDir DestinationFolder$ + "WIM_x64"
MkDir DestinationFolder$ + "WIM_x86"

' Export all the x64 and x86 editions to the WIM_x64 and WIM_x86 folders.

Cls
Print
Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Exporting All Windows Editions"
Print "[             ] Creating Base Image"
Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
Print "[             ] Creating Final ISO Image"

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        MountISO Temp$
        For y = 1 To IndexCount(x)
            Select Case FileSourceType$(x)
                Case "x64"
                    SourceArcFlag$ = ""
                    DestArcFlag$ = "WIM_x64"
                Case "x86"
                    SourceArcFlag$ = ""
                    DestArcFlag$ = "WIM_x86"
                Case "x64_DUAL"
                    SourceArcFlag$ = "\x64"
                    DestArcFlag$ = "WIM_x64"
                Case "x86_DUAL"
                    SourceArcFlag$ = "\x86"
                    DestArcFlag$ = "WIM_x86"
            End Select
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
Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
Print "[             ] Creating Final ISO Image"

If ((x64UpdateImageCount > 0) And (x86UpdateImageCount > 0)) Then
    AllFilesAreSameArc = 0
    SingleImageTag$ = ""
Else
    AllFilesAreSameArc = 1
    If (x64UpdateImageCount > 0) Then
        SingleImageTag$ = "\x64"
    ElseIf (x86UpdateImageCount > 0) Then
        SingleImageTag$ = "\x86"
    End If
End If

' To ensure that DestinationFolder$ is always specified consistently without a trailing backslash, we will
' run it through the CleanPath routine.

CleanPath DestinationFolder$
DestinationFolder$ = Temp$

' Jump to the routine for creating a base image. Here we determine whether we need to run code for a Single or Dual Architecture project

Select Case ProjectType$
    Case "x64", "x86"
        GoTo MMBIProjectIsSingleArchitecture
    Case "DUAL"
        GoTo MMBIProjectIsDualArchitecture
End Select

MMBIProjectIsSingleArchitecture:

For x = 1 To TotalFiles
    If UpdateFlag$(x) = "Y" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        MountISO Temp$

        ' If an x64 folder exists, then even though the project is a single architecture type project, the source is a dual architecture source.
        ' This means that we need to copy the contents of the x64 or x86 folder to the root and not to the x64 or x86 folder.

        If _DirExists(MountedImageDriveLetter$ + "\x64") Then
            Select Case ExcludeAutounattend$
                Case "Y"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + SingleImageTag$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
                    + CHR$(34) + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd autounattend.xml > NUL"
                Case "N"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + SingleImageTag$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
                    + CHR$(34) + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd > NUL"
            End Select
            Shell Cmd$
        Else
            Select Case ExcludeAutounattend$
                Case "Y"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
                    + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd autounattend.xml > NUL"
                Case "N"
                    Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
                    + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd > NUL"
            End Select
            Shell Cmd$
        End If
        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Exit For
    End If
Next x

' When we arrive here, the base image for the single architecure type project has been completed.

GoTo MMBIDoneCreatingBaseImage

MMBIProjectIsDualArchitecture:

If DualArcImagePath$ = "" Then GoTo MMBINoDualImageDistrib

MountISO DualArcImagePath$
Select Case ExcludeAutounattend$
    Case "Y"
        Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + " " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
        + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd autounattend.xml > NUL"
    Case "N"
        Cmd$ = "robocopy " + CHR$(34) + MountedImageDriveLetter$ + " " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files" + CHR$(34)_
        + " /mir /nfl /ndl /njh /njs /a-:rsh /xf install.wim install.esd > NUL"
End Select
Shell Cmd$
Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + DualArcImagePath$ + "'" + Chr$(34) + Chr$(34)
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
GoTo MMBIDualArcBaseImageDone

MMBINoDualImageDistrib:

' If no dual architecture image is available, copy the needed files from x64 and x86 media
' and dynamically generate the bcd files needed.

For x = 1 To TotalFiles
    If FileSourceType$(x) = "x64" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        MountISO Temp$
        Select Case ExcludeAutounattend$
            Case "Y"
                Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xd sources support /xf bcd install.wim autounattend.xml /a-:rsh > NUL"
            Case "N"
                Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xd sources support /xf bcd install.wim /a-:rsh > NUL"
        End Select
        Shell _Hide Cmd$
        Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x64" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xf install.wim autounattend.xml /a-:rsh > NUL"
        Shell _Hide Cmd$
        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-Diskimage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        GoTo MMBI_Do_x86_Build_Base
    End If
Next x

MMBI_Do_x86_Build_Base:

For x = 1 To TotalFiles
    If FileSourceType$(x) = "x86" Then
        Temp$ = FolderArray$(x) + FileArray$(x)
        MountISO Temp$
        Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files" + Chr$(34) + " autorun.inf setup.exe /a-:rsh > NUL"
        Shell _Hide Cmd$
        Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + "\efi\boot" + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\efi\boot" + Chr$(34) + " bootia32.efi /a-:rsh > NUL"
        Shell _Hide Cmd$
        Cmd$ = "robocopy " + Chr$(34) + MountedImageDriveLetter$ + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x86" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xf install.wim autounattend.xml /a-:rsh > NUL"
        Shell _Hide Cmd$
        Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-Diskimage " + Chr$(34) + "'" + Temp$ + "'" + Chr$(34) + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        GoTo MMBI_Build_Bcd_Files
    End If
Next x

MMBI_Build_Bcd_Files:

GoSub Create_Reg_Files

' Create a template bcd hive for each of the two locations where we need this file

Cmd$ = "bcdedit /createstore " + Chr$(34) + DestinationFolder$ + "\ISO_Files\boot\bcd" + Chr$(34)
Shell _Hide Cmd$
Cmd$ = "bcdedit /createstore " + Chr$(34) + DestinationFolder$ + "\ISO_Files\efi\microsoft\boot\bcd" + Chr$(34)
Shell _Hide Cmd$

' Load the template hives into the registry

Cmd$ = "reg load HKLM\BCD_BIOS " + Chr$(34) + DestinationFolder$ + "\ISO_Files\boot\bcd" + Chr$(34)
Shell _Hide Cmd$
Cmd$ = "reg load HKLM\BCD_EFI " + Chr$(34) + DestinationFolder$ + "\ISO_Files\efi\microsoft\boot\bcd" + Chr$(34)
Shell _Hide Cmd$

' Create a "permissions.txt" file that defines registry permission changes to be applied by regini

Temp$ = TempLocation$ + "\permissions.txt"
Open Temp$ For Output As #1
Print #1, "\Registry\machine\BCD_BIOS [1 6 17]"
Print #1, "\Registry\machine\BCD_BIOS\Description [1 6 17]"
Print #1, "\Registry\machine\BCD_BIOS\Objects [1 6 17]"
Print #1, "\Registry\machine\BCD_EFI [1 6 17]"
Print #1, "\Registry\machine\BCD_EFI\Description [1 6 17]"
Print #1, "\Registry\machine\BCD_EFI\Objects [1 6 17]"
Close #1

' Run regini to alter the permissions

Cmd$ = "regini " + Chr$(34) + TempLocation$ + "\permissions.txt" + Chr$(34)
Shell _Hide Cmd$
Temp$ = TempLocation$ + "\permissions.txt"
Kill Temp$

' Import the registry files in order to apply those settings to out template hives

Cmd$ = "reg import " + Chr$(34) + TempLocation$ + "\bcd_bios.reg" + Chr$(34)
Shell _Hide Cmd$
Cmd$ = "reg import " + Chr$(34) + TempLocation$ + "\bcd_efi.reg" + Chr$(34)
Shell _Hide Cmd$

' Unload the registry hives, committing the changes to what were the templates, making them
' the final bcd hive files.

Cmd$ = "reg unload HKLM\BCD_BIOS"
Shell _Hide Cmd$
Cmd$ = "reg unload HKLM\BCD_EFI"
Shell _Hide Cmd$

' Delete the temporary registry files used to create the bcd files

Temp$ = TempLocation$ + "\bcd_bios.reg"
Kill Temp$
Temp$ = TempLocation$ + "\bcd_efi.reg"
Kill Temp$

MMBIDualArcBaseImageDone:

' When we arrive here, the creation of the dual architecture base image is completed

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

    Temp$ = DestinationFolder$ + "\ISO_Files\x64\sources"

    If _DirExists(Temp$) Then
        If Not (_FileExists(Temp$ + "\ei.cfg")) Then
            Open (Temp$ + "\ei.cfg") For Output As #1
            Print #1, "[CHANNEL]"
            Print #1, "Retail"
            Close #1
        End If
    End If

    Temp$ = DestinationFolder$ + "\ISO_Files\x86\sources"

    If _DirExists(Temp$) Then
        If Not (_FileExists(Temp$ + "\ei.cfg")) Then
            Open (Temp$ + "\ei.cfg") For Output As #1
            Print #1, "[CHANNEL]"
            Print #1, "Retail"
            Close #1
        End If
    End If
End If

MMBIDoneCreatingBaseImage:

' Moving the updated install.wim file(s) to the base image

Cls
Print
Print "[  COMPLETED  ] Exporting All Windows Editions"
Print "[  COMPLETED  ] Creating Base Image"
Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Moving Updated WIM Files to Base Image and Syncing File Versions"
Print "[             ] Creating Final ISO Image"

If ProjectType$ = "DUAL" Then
    If AllFilesAreSameArc = 1 Then
        Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\WIM_x64\install.wim " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
        + "\Sources" + CHR$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "move /Y " + CHR$(34) + DestinationFolder$ + "\WIM_x86\install.wim " + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\ISO_Files"_
        + "\Sources" + CHR$(34) + " > NUL"
        Shell _Hide Cmd$
    Else
        Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "\WIM_x64\install.wim " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x64\Sources" + Chr$(34) + " > NUL"
        Shell _Hide Cmd$
        Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "\WIM_x86\install.wim " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\x86\Sources" + Chr$(34) + " > NUL"
        Shell _Hide Cmd$
    End If
End If

If ProjectType$ = "SINGLE" Then
    If x64UpdateImageCount > 0 Then
        Temp$ = DestinationFolder$ + "\WIM_x64\install.wim"
        If _FileExists(Temp$) Then
            Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
            Shell _Hide Cmd$
            Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "\WIM_x64\install.wim " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
            Shell _Hide Cmd$
        End If
    End If
    If x86UpdateImageCount > 0 Then
        Temp$ = DestinationFolder$ + "\WIM_x86\install.wim"
        If _FileExists(Temp$) Then
            Cmd$ = "md " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
            Shell _Hide Cmd$
            Cmd$ = "move /Y " + Chr$(34) + DestinationFolder$ + "\WIM_x86\install.wim " + Chr$(34) + " " + Chr$(34) + DestinationFolder$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
            Shell _Hide Cmd$
        End If
    End If
End If

FinalImageName$ = DestinationFolder$ + "\" + UserSelectedImageName$

' Technical Note: OSCDIMG does not hide its output by simply redirecting to NUL. By using " > NUL 2>&1" we work around this.
' How this works: Standard output is going to NUL and standard error output (file descriptor 2) is being sent to standard output
' (file descriptor 1) so both error and normal output go to the same place.

Cls
Print
Print "[  COMPLETED  ] Exporting All Windows Editions"
Print "[  COMPLETED  ] Creating Base Image"
Print "[  COMPLETED  ] Moving Updated WIM Files to Base Image and Syncing File Versions"
Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Final ISO Image"

' Clear the read-only, system, and hidden attributes from all source files

Cmd$ = "attrib -h -s -r " + Chr$(34) + DestinationFolder$ + "\ISO_Files\*.*" + Chr$(34) + " /s /d"
Shell _Hide Chr$(34) + Cmd$ + Chr$(34)

' Create the final ISO image file

Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -m -o -u2 -udfver102 -bootdata:2#p0,e,b" + CHR$(34) + DestinationFolder$ + "\ISO_Files\boot\etfsboot.com"_
+ CHR$(34) + "#pEF,e,b" + CHR$(34) + DestinationFolder$ + "\ISO_Files\efi\microsoft\boot\efisys.bin" + CHR$(34) + " " + CHR$(34) + DestinationFolder$_
+ "\ISO_Files" + CHR$(34) + " " + CHR$(34) + FinalImageName$ + CHR$(34) + " > NUL 2>&1"
Shell Chr$(34) + Cmd$ + Chr$(34)

Cls
Print
Print "[  COMPLETED  ] Exporting All Windows Editions"
Print "[  COMPLETED  ] Creating Base Image"
Print "[  COMPLETED  ] Moving Updated WIM Files to Base Image and Syncing File Versions"
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
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "\WIM_x86" + Chr$(34) + " /s /q"
Shell _Hide Cmd$

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
    Print "Enter the path to the folder with the Windows files to make into an ISO image: ";: Input "", MakeBootablePath$

Loop While MakeBootablePath$ = ""

CleanPath MakeBootablePath$
MakeBootablePath$ = Temp$
TempPath$ = MakeBootablePath$ + "\sources\install.wim"

' We cannot check for all files, but we are at least checking to see if an "INSTALL.WIM" is present at the specified location as a simple
' sanity check that the folder specified is likely valid.

If Not ((_FileExists(TempPath$)) Or (_FileExists(MakeBootablePath$ + "\x64\sources\install.wim")) Or (_FileExists(MakeBootablePath$ + "\x86\sources\install.wim"))) Then
    Print
    Color 14, 4: Print "That path is not valid.";: Color 15: Print " No INSTALL.WIM file found at that location. Please try again."
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
    Print "Enter the destination path. This is the path only ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " a file name: ";: Input "", DestinationFolder$
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
    Print "Enter the name of the file to create, ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension: ";: Input "", DestinationFileName$
Loop While DestinationFileName$ = ""

GetVolumeName1:

' Get the volume name for the ISO image

Cls
Input "Enter the volume name to give the ISO image or press Enter for none: ", VolumeName$

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

Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -m -o -u2 -udfver102 -l" + CHR$(34) + VolumeName$ + CHR$(34) + " -bootdata:2#p0,e,b" + CHR$(34)_
+ MakeBootablePath$ + "\boot\etfsboot.com" + CHR$(34) + "#pEF,e,b" + CHR$(34) + MakeBootablePath$ + "\efi\microsoft\boot\efisys.bin" + CHR$(34)_
+ " " + CHR$(34) + MakeBootablePath$ + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "\" + DestinationFileName$ + ".iso" + CHR$(34) + " > NUL 2>&1"
Print "Creating the ISO image. Please standby..."
Shell Chr$(34) + Cmd$ + Chr$(34)
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
' in the order that you want. You can also remove Windows editions from the image. Finally, if you are working with a dual
' architecture image and you choose to keep only editions of one architecture type, this routine will convert the resulting
' image into a single architecture image.

' Ask for source image file. Verify that it is a valid image.

GetImageToReorg:

' Initialize variable
SourceImage$ = ""

Cls
Input "Please enter the full path and file name of the image to reorganize: ", SourceImage$
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
    Case "x64", "x86"
        FileSourceType$ = "SINGLE"
    Case "DUAL"
        FileSourceType$ = "DUAL"
    Case "NONE"
        Cls
        Color 14, 4: Print "The image specified is not valid.";: Color 15: Print " Please specify a valid image."
        Pause
        GoTo ChangeOrder
End Select

'Init variables

x64_IndexOrder$ = ""
x86_IndexOrder$ = ""
IndexOrder$ = ""
x64ImageCount = 0
x86ImageCount = 0
SingleImageCount = 0

' Ask the user for the list of indices in the order in which they want them arranged. When getting the indices for a single architecture type, the user must enter something.
' An empty string is not valid so don't allow this. For a dual architecture image, it is permissible for a user to specify no indices for either the x64 or x86 editions so
' we will allow that here. However, at least one index from one of those architecture types is needed so a little later we will check to make sure that the user supplied
' at one index for either x64 or x86.

GetTheIndices:

If FileSourceType$ = "DUAL" Then
    Reorgx86:
    Cls
    Print "The file that you specified is a dual architecture image file."
    Print "Enter the ";: Color 0, 14: Print ">> x86 <<";: Color 15: Print " index order, ENTER for none, or HELP: ";
    Input "", x86_IndexOrder$
    If x86_IndexOrder$ = "" GoTo Reorgx64
    If UCase$(x86_IndexOrder$) = "HELP" Then
        Cls
        Print "You can enter a single index number or multiple index numbers. To enter a contiguous range of index numbers,"
        Print "separate the numbers with a dash like this: 1-4. For non contiguous indices, separate them with a space like"
        Print "this: 1 3. You can also combine both methods like this: 1-3 5 7-9. Numbers can be entered in both ascending"
        Print "and descending order like this: 1 3 9-7 5."
        Pause
        GoTo Reorgx86
    End If

    ProcessRangeOfNums x86_IndexOrder$, 0
    If ValidRange = 0 Then
        Cls
        Color 14, 4: Print "You did not enter valid values.";: Color 15: Print " Enter a valid set of values.": Color 15
        GoTo Reorgx86
    End If
    x86ImageCount = TotalNumsInArray
    ReDim x86Array(x86ImageCount) As Integer
    For x = 1 To x86ImageCount
        x86Array(x) = RangeArray(x)
    Next x

    Reorgx64:

    Cls
    Print "The file that you specified is a dual architecture image file."
    Print "Enter the ";: Color 0, 14: Print ">> x64 <<";: Color 15: Print " index order, ENTER for none, or HELP: ";
    Input "", x64_IndexOrder$
    If x64_IndexOrder$ = "" GoTo End_Reorgx64
    If UCase$(x64_IndexOrder$) = "HELP" Then
        Cls
        Print "You can enter a single index number or multiple index numbers. To enter a contiguous range of index numbers,"
        Print "separate the numbers with a dash like this: 1-4. For non contiguous indices, separate them with a space like"
        Print "this: 1 3. You can also combine both methods like this: 1-3 5 7-9. Numbers can be entered in both ascending"
        Print "and descending order like this: 1 3 9-7 5."
        Pause
        GoTo Reorgx64
    End If

    ProcessRangeOfNums x64_IndexOrder$, 0
    If ValidRange = 0 Then
        Cls
        Color 14, 4: Print "You did not enter valid values.";: Color 15: Print " Enter a valid set of values.": Color 15
        GoTo Reorgx64
    End If
    x64ImageCount = TotalNumsInArray
    ReDim x64Array(x64ImageCount) As Integer
    For x = 1 To x64ImageCount
        x64Array(x) = RangeArray(x)
    Next x

    End_Reorgx64:

Else

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
End If

' Verify that the index numbers specified are valid.

' We will now get WIM info and save it to a file called WIM_Info.txt. We will parse that file to verify that the indices
' selected are valid. If not, we will ask the user to choose a valid index.

Print
Print "Verifying that the index number(s) supplied are valid."
Print
Print "Please standby..."
Print

' If the source is a dual architecture image, make sure that the user has specified at least one index value from one of the architecture types.

If (FileSourceType = "DUAL") And (x64_IndexOrder = "") And (x86_IndexOrder = "") Then
    Cls
    Color 14, 4: Print "Invalid Entry!": Color 15
    Print
    Print "The source is a dual architecture image but you have not selected any indices from either the x64 or the x86"
    Print "editions available in this image."
    Print
    Print "Please specify at least one valid index."
    Pause
    GoTo GetTheIndices
End If

' We are now going to get information regarding the WIM file(s) on the source. We will parse that information to determine what the highest numbered index is
' for a single architecture image or for the x64 editions and x86 editions on a dual architecture image. This will allow us to check that no index numbers
' higher than valid have been supplied.

GetWimInfo_Main SourceImage$, 1

' Initialize variables

Highest_Single = 0
Highest_x86 = 0
Highest_x64 = 0

Open "WIM_Info.txt" For Input As #1

Select Case FileSourceType$
    Case "DUAL"
        Do
            Line Input #1, ReadLine$
            If InStr(ReadLine$, "Index :") Then
                Temp$ = (Right$(ReadLine$, (Len(ReadLine$) - _InStrRev(ReadLine$, ":"))))
            End If
        Loop Until InStr(ReadLine$, "x64 Editions")
        Highest_x86 = Val(Temp$)
        Do
            Line Input #1, ReadLine$
            If InStr(ReadLine$, "Index :") Then
                Temp$ = (Right$(ReadLine$, (Len(ReadLine$) - _InStrRev(ReadLine$, ":"))))
            End If
        Loop Until EOF(1)
        Highest_x64 = Val(Temp$)
    Case "SINGLE"
        Do
            Line Input #1, ReadLine$
            If InStr(ReadLine$, "Index :") Then
                Temp$ = (Right$(ReadLine$, (Len(ReadLine$) - _InStrRev(ReadLine$, ":"))))
            End If
        Loop Until EOF(1)
        Highest_Single = Val(Temp$)
End Select

' Close and delete the WIM_Info.txt file since we are now done using it

Close #1
Kill "WIM_Info.txt"

' Initialize variable

ValidRange = 1

Select Case FileSourceType$
    Case "DUAL"
        If x86ImageCount > 0 Then
            For x = 1 To x86ImageCount
                If x86Array(x) > Highest_x86 Then ValidRange = 0
            Next x
        End If
        If x64ImageCount > 0 Then
            For x = 1 To x64ImageCount
                If x64Array(x) > Highest_x64 Then ValidRange = 0
            Next x
        End If
    Case "SINGLE"
        For x = 1 To SingleImageCount
            If SingleArray(x) > Highest_Single Then ValidRange = 0
        Next x
End Select

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
Input "Enter the destination for the project (path only, no file name or extension): ", Destination$
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

If _DirExists(Destination$ + "\x64") Then GoTo ReorgAssetExists
If _DirExists(Chr$(34) + Destination$ + "\x86" + Chr$(34)) Then GoTo ReorgAssetExists
If _DirExists(Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34)) Then GoTo ReorgAssetExists
If _FileExists(Chr$(34) + Destination$ + "\install.wim" + Chr$(34)) Then GoTo ReorgAssetExists
GoTo ReorgCleanup

ReorgAssetExists:

Cls
Print "This routine creates the temporary folders named x64, x86, ISO_Files, and a file named install.wim"
Print "in the destination folder. At least one of these already exists there."
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
Shell _Hide "rmdir " + Chr$(34) + Destination$ + "\x64" + Chr$(34) + " /s /q"
Shell _Hide "rmdir " + Chr$(34) + Destination$ + "\x86" + Chr$(34) + " /s /q"
Shell _Hide "rmdir " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34) + " /s /q"
Shell _Hide "del " + Chr$(34) + Destination$ + "\install.wim" + Chr$(34) + " /s /q"
Shell _Hide "md " + Chr$(34) + Destination$ + "\x64" + Chr$(34)
Shell _Hide "md " + Chr$(34) + Destination$ + "\x86" + Chr$(34)
Shell _Hide "md " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34)

' Mount the ISO image and copy the editions to our working folder in the proper order

Cls
Print "Exporting Windows editions..."
MountISO SourceImage$
ImageSourceDrive$ = MountedImageDriveLetter$

If FileSourceType$ = "DUAL" Then
    If (x64ImageCount > 0) And (x86ImageCount > 0) Then
        SRC$ = ImageSourceDrive$ + "\x64\Sources\install.wim"
        DST$ = Destination$ + "\x64\install.wim"
        For x = 1 To x64ImageCount
            Locate 3, 1: Print "Exporting x64 image"; x; "of"; x64ImageCount
            IDX$ = LTrim$(Str$(x64Array(x)))
            Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Export-Image /SourceImageFile:" + SRC$ + " /SourceIndex:" + IDX$ + " /DestinationImageFile:" + Chr$(34) + DST$ + Chr$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Next x
        SRC$ = ImageSourceDrive$ + "\x86\Sources\install.wim"
        DST$ = Destination$ + "\x86\install.wim"
        For x = 1 To x86ImageCount
            Locate 3, 1: Print "Exporting x86 image"; x; "of"; x86ImageCount
            IDX$ = LTrim$(Str$(x86Array(x)))
            Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Export-Image /SourceImageFile:" + SRC$ + " /SourceIndex:" + IDX$ + " /DestinationImageFile:" + Chr$(34) + DST$ + Chr$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Next x
    End If
    If (x64ImageCount > 0) And (x86ImageCount = 0) Then
        SRC$ = ImageSourceDrive$ + "\x64\Sources\install.wim"
        DST$ = Destination$ + "\x64\install.wim"
        For x = 1 To x64ImageCount
            Locate 3, 1: Print "Exporting x64 image"; x; "of"; x64ImageCount
            IDX$ = LTrim$(Str$(x64Array(x)))
            Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Export-Image /SourceImageFile:" + SRC$ + " /SourceIndex:" + IDX$ + " /DestinationImageFile:" + Chr$(34) + DST$ + Chr$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Next x
    End If
    If (x86ImageCount > 0) And (x64ImageCount = 0) Then
        SRC$ = ImageSourceDrive$ + "\x86\Sources\install.wim"
        DST$ = Destination$ + "\x86\install.wim"
        For x = 1 To x86ImageCount
            Locate 3, 1: Print "Exporting x86 image"; x; "of"; x86ImageCount
            IDX$ = LTrim$(Str$(x86Array(x)))
            Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Export-Image /SourceImageFile:" + SRC$ + " /SourceIndex:" + IDX$ + " /DestinationImageFile:" + Chr$(34) + DST$ + Chr$(34)
            Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
        Next x
    End If
End If

If FileSourceType$ = "SINGLE" Then
    SRC$ = ImageSourceDrive$ + "\Sources\install.wim"
    DST$ = Destination$ + "\install.wim"
    For x = 1 To SingleImageCount
        Locate 3, 1: Print "Exporting image"; x; "of"; SingleImageCount
        IDX$ = LTrim$(Str$(SingleArray(x)))
        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Export-Image /SourceImageFile:" + SRC$ + " /SourceIndex:" + IDX$ + " /DestinationImageFile:" + Chr$(34) + DST$ + Chr$(34)
        Shell _Hide Chr$(34) + Cmd$ + Chr$(34)
    Next x
End If

Cls
Print "Reorganizing the image in the order you specified."

' Copy the files needed to create the base image. If the source file was a dual architecture image but the user is only keeping editions from
' one architecture type, then we will reorganize the new image directory structure to make it a single architecture image.

If (FileSourceType$ = "SINGLE") Or ((FileSourceType$ = "DUAL") And (x64ImageCount > 0) And (x86ImageCount > 0)) Then
    Cmd$ = "robocopy " + Chr$(34) + ImageSourceDrive$ + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xf install.wim"
    Shell Cmd$
End If

If (FileSourceType$ = "DUAL") And (x64ImageCount > 0) And (x86ImageCount = 0) Then
    Cmd$ = "robocopy " + Chr$(34) + ImageSourceDrive$ + "\x64" + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xf install.wim"
    Shell Cmd$
End If

If (FileSourceType$ = "DUAL") And (x86ImageCount > 0) And (x64ImageCount = 0) Then
    Cmd$ = "robocopy " + Chr$(34) + ImageSourceDrive$ + "\x86" + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34) + " /mir /nfl /ndl /njh /njs /xf install.wim"
    Shell Cmd$
End If

If _FileExists(Destination$ + "\install.wim") Then
    Cmd$ = "move /Y " + Chr$(34) + Destination$ + "\install.wim" + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
    Shell _Hide Cmd$
End If

If _FileExists(Destination$ + "\x64\install.wim") Then
    If x86ImageCount > 0 Then
        Cmd$ = "move /Y " + Chr$(34) + Destination$ + "\x64\install.wim" + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files\x64\Sources" + Chr$(34) + " > NUL"
        Shell _Hide Cmd$
    Else
        Cmd$ = "move /Y " + Chr$(34) + Destination$ + "\x64\install.wim" + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
        Shell _Hide Cmd$
    End If
End If

If _FileExists(Destination$ + "\x86\install.wim") Then
    If x64ImageCount > 0 Then
        Cmd$ = "move /Y " + Chr$(34) + Destination$ + "\x86\install.wim" + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files\x86\Sources" + Chr$(34) + " > NUL"
        Shell _Hide Cmd$
    Else
        Cmd$ = "move /Y " + Chr$(34) + Destination$ + "\x86\install.wim" + Chr$(34) + " " + Chr$(34) + Destination$ + "\ISO_Files\Sources" + Chr$(34) + " > NUL"
        Shell _Hide Cmd$
    End If
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
Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -m -o -u2 -udfver102 -l" + CHR$(34) + VolumeLabel$ + CHR$(34) + " -bootdata:2#p0,e,b" + CHR$(34)_
+ Destination$ + "\ISO_Files\boot\etfsboot.com" + CHR$(34) + "#pEF,e,b" + CHR$(34) + Destination$ + "\ISO_Files\efi\microsoft\boot\efisys.bin"_
+ CHR$(34) + " " + CHR$(34) + Destination$ + "\ISO_Files" + CHR$(34) + " " + CHR$(34) + Destination$ + "\" + ReorgFileName$ + CHR$(34) + " > NUL 2>&1"
Shell Chr$(34) + Cmd$ + Chr$(34)


' Project is done. Cleanup the temporary files.

Cls
Print "Removing the temporary files used to create the new image..."
Shell _Hide "rmdir " + Chr$(34) + Destination$ + "\x64" + Chr$(34) + " /s /q"
Shell _Hide "rmdir " + Chr$(34) + Destination$ + "\x86" + Chr$(34) + " /s /q"
Shell _Hide "rmdir " + Chr$(34) + Destination$ + "\ISO_Files" + Chr$(34) + " /s /q"
Shell _Hide "del " + Chr$(34) + Destination$ + "\install.wim" + Chr$(34) + " /s /q"

' Inform the user that the project is done and then return to the main menu.

Cls
Print "Done!"
Print
Print "The updated file can be found here:"
Print
Color 10: Print Destination$: Color 15
Print
Print "The file name is the same as the original file: ";: Color 10: Print ReorgFileName$: Color 15

If ((x64ImageCount > 0) And (x86ImageCount = 0)) Or ((x86ImageCount > 0) And (x64ImageCount = 0)) Then
    Print
    Color 0, 10: Print "NOTE:";: Color 15: Print " The original source image was a dual architecture image but you have kept only editions from one"
    Print "architecture type, so we have converted the new image into a single architecture image."
    Print
End If

Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' ******************************************************************
' * Get WIM info - display basic info for each WIM in an ISO image *
' ******************************************************************

GetWimInfo:

Do
    Cls
    Print "Enter the full path to the ISO image from which to get information."
    Input "Include the file name and extension: ", SourcePath$
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

' Display WIM_Info.txt which lists all of the indicies, then delete it if user no longer needs it.

Cls
DisplayFile "WIM_Info.txt"

AskToSaveWimInfo1:
Cls
Print "A copy of the information just displayed can be saved to a file named WIM_Info.txt in the same location"
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
    Kill "WIM_Info.txt"
End If

If Temp$ = "Y" Then
    Shell Chr$(34) + "move WIM_Info.txt " + Chr$(34) + ProgramStartDir$ + Chr$(34) + " > NUL" + Chr$(34)
    Cls
    Print "The file has been saved as:"
    Print
    Color 10: Print ProgramStartDir$; "\WIM_Info.txt": Color 15
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
    Input "Include the file name and extension: ", SourcePath$

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
Print "Checking to see if the image specified is dual or single architecture."
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

Do
    Cls
    Print "Note that you can save your updated image to the same location where the original is located."
    Print "You can even use the same file name if you want to update that file in its current location."
    Print
    Input "Enter the destination path without a file name or extension: ", DestinationFolder$
Loop While DestinationFolder$ = ""

CleanPath DestinationFolder$
DestinationFolder$ = Temp$ + "\"

Do
    Cls
    Print "Enter the name of the ISO image file to create ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension: ";: Input "", OutputFileName$
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

' If we are dealing with a dual architecture image, then the user needs to specify whether to modify
' one of the entries in the x86 or the x64 WIM. We will set ArchitectureChoice$ to either x86 or x64.
' If the image is NOT dual architecture, then we set ArchitectureChoice$ to "".

If ProjectArchitecture$ = "DUAL" Then
    Do
        Cls
        Input "Do you want to modify the information for an x86 or x64 entry (type x86 or x64): ", ArchitectureChoice$
    Loop Until ((ArchitectureChoice$ = "x86") Or (ArchitectureChoice$ = "x64"))
Else
    ArchitectureChoice$ = ""
End If

If ArchitectureChoice$ <> "" Then
    ArchitectureChoice$ = "\" + ArchitectureChoice$
End If

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
        Print "A copy of the information just displayed can be saved to a file named WIM_Info.txt in the same location"
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
            Kill "WIM_Info.txt"
        End If
        If Temp$ = "Y" Then
            Shell Chr$(34) + "move WIM_Info.txt " + Chr$(34) + ProgramStartDir$ + Chr$(34) + " > NUL" + Chr$(34)
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
+ " " + IndexString$ + " " + CHR$(34) + EditionName$ + CHR$(34) + " " + CHR$(34) + Description$ + CHR$(34) + " /check > NUL"
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
Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " -m -o -u2 -udfver102 -l" + CHR$(34) + VolumeName$ + CHR$(34) + " -bootdata:2#p0,e,b" + CHR$(34)_
+ DestinationFolder$ + "ISO_Files\boot\etfsboot.com" + CHR$(34) + "#pEF,e,b" + CHR$(34) + DestinationFolder$ + "ISO_Files\efi\microsoft\boot\efisys.bin"_
+ CHR$(34) + " " + CHR$(34) + DestinationFolder$ + "ISO_Files" + CHR$(34) + " " + CHR$(34) + DestinationFolder$ + OutputFileName$ + ".ISO" + CHR$(34) + " > NUL 2>&1"
Shell Chr$(34) + Cmd$ + Chr$(34)
Print "*******************************"
Print "* Cleaning up temporary files *"
Print "*******************************"
Print
Cmd$ = "rmdir " + Chr$(34) + DestinationFolder$ + "ISO_Files" + Chr$(34) + " /s /q"
Shell _Hide Cmd$
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
    Input "Enter the full path to the location where you want the drivers to be exported: ", ExportFolder$
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
Print #ff, "@echo off"
Print #ff, ""
Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ":: Change to the directory where the batch file is run from ::"
Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ""
Print #ff, "cd /d %~dp0"
Print #ff, ""
Print #ff, "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
Print #ff, ":: Check to see if this batch file is being run as Administrator. If it is not, then rerun the batch file ::"
Print #ff, ":: automatically as admin and terminate the intial instance of the batch file.                            ::"
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
Print #ff, "echo Drivers are being installed."
Print #ff, "echo Please be patient since this process may take a while..."
Print #ff, ""
Print #ff, "pnputil /add-driver *.inf /subdirs /install > NUL"
Print #ff, ""
Print #ff, "cls"
Print #ff, "echo Drivers have been installed. Please be aware that a reboot may be needed."
Print #ff, "echo."
Print #ff, "pause"
Close #ff
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
    Input "Enter the path to the drivers that are in .CAB files: ", SourceFolder$
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
    Input "Enter the destination path: ", DestinationFolder$
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
    Input "Please enter path: ", VHDXPath$
Loop While VHDXPath$ = ""

' Remove quotes and trailing backslash.

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
    Input "Enter file name: ", VHDXFileName$
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
Print "We need to know what ISO image contains the edition of Windows that you want to install to the VHD. Please provide"
Print "the full path, ";: Color 0, 10: Print "including the file name";: Color 15: Print ", to this ISO image."
Print
Input "Enter the full path to the ISO image: ", SourceImage$
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
    Print "Enter the full path to the location where we should save the VHD ";: Color 0, 10: Print "not including a filename";: Color 15: Print "."
    Print
    Input "Enter the path: ", Destination$
Loop While Destination$ = ""

CleanPath Destination$
DestinationFolder$ = Temp$ + "\"

' Check to see if the destination specified is on a removable disk

Cls
Print "Performing a check to see if the destination you specified is a removable disk."
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
    Case 1
        Cls
        Print "The disk that you specified is a removable disk. ";: Color 14, 4: Print "Please specify a fixed disk.": Color 15
        Pause
        GoTo GetVHDDestination
    Case 0
        ' if the returned value was a 0, no action is necessary. The program will continue normally.
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
    Input "Enter the filename: ", VHDFilename$
Loop While VHDFilename$ = ""

' Build the full path including the filename

Destination$ = Destination$ + VHDFilename$ + ".VHDX"

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

SpecifyDriveLetter:

Cls
Print "Enter the drive letter to assign to the VHD. NOTE: We only need this drive letter on a temporary basis while we"
Print "deploy Windows to it. After you reboot, the drive letter will no longer be in use."
Print
Input "Enter the drive letter (Letter only, no colon): "; DriveLetter$
DriveLetter$ = UCase$(DriveLetter$)

If (Len(DriveLetter$) > 1) Or (DriveLetter$) = "" Or ((Asc(DriveLetter$)) < 65) Or ((Asc(DriveLetter$)) > 90) Then
    Print
    Color 14, 4: Print "That was not a valid entry.";: Color 15: Print " Please try again."
    Print
    GoTo SpecifyDriveLetter
End If

If _DirExists(DriveLetter$ + ":") Then
    Print
    Color 14, 4: Print "That drive letter is already in use.": Color 15
    Pause
    GoTo SpecifyDriveLetter
End If

Cls
Print "Enter the description to be displayed in the boot menu. Example: Win 10 Pro (VHD)"
Print
Input "Enter description: ", Description$
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
Print "that you want to apply to the VHD, then cancel out of this routine and run the option "; Chr$(34); "Get WIM info - display basic"
Print "info for each WIM in an ISO image"; Chr$(34); " from the main menu."
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
Print "Standby while we create and mount the VHD to the drive letter "; UCase$(DriveLetter$); ":"

Select Case VHD_Type
    Case 1
        GoTo CreateMBR_VHD
    Case 2
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
Print #1, "echo shrink minimum=500"
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
' to prevent the following changes bcdboot bcdedit commands from causing the system
' to ask for the recovery key

' If BitLocker is enabled we are suspending temporarily (until the next boot into this instace of Windows)

Cmd$ = "manage-bde -protectors -disable C: > NUL"
Shell Cmd$
Print "Updating boot information"
Cmd$ = "bcdboot " + DriveLetter$ + ":\Windows > NUL"
Shell Cmd$
Cmd$ = "bcdedit /set {default} description " + Chr$(34) + Description$ + Chr$(34) + " > NUL"
Shell Cmd$
Print
Print "***********************************************"
Print "* Done. Windows has been deployed to VHD file *"
Print "* and the host boot menu has been updated.    *"
Print "***********************************************"
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


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

    Input "Enter the path containing the data to place into an ISO image: ", SourcePath$
Loop While SourcePath$ = ""

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
    Input "Enter the destination path. This is the path only without a file name: ", DestinationPath$
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
    Print "Enter the name of the file to create, ";: Color 0, 10: Print "WITHOUT";: Color 15: Print " an extension: ";: Input "", DestinationFileName$
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

Cmd$ = CHR$(34) + OSCDIMGLocation$ + CHR$(34) + " " + "-o -m -h -k -u2 -udfver102 -l" + CHR$(34) + VolumeName$ + CHR$(34) + " " + CHR$(34) + SourcePath$_
+ CHR$(34) + " " + CHR$(34) + DestinationPathAndFile$ + CHR$(34) + " > NUL 2>&1"

' Create the ISO image

Cls
Print "Creating the image. Please standby..."
Shell Chr$(34) + Cmd$ + Chr$(34)
Print
Print "ISO Image created."
Pause
ChDir ProgramStartDir$: GoTo BeginProgram


' *****************************
' * Cleanup files and folders *
' *****************************

GetFolderToClean:

DestinationPath$ = "" ' Set initial value

Do
    Cls
    Input "Please enter the full path to the project folder to be cleaned: ", DestinationPath$
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
Color 0, 13
Print "    1) Get general help on the use of this program                                                              "
Color 0, 14
Print "    2) Inject Windows updates into one or more Windows editions and create a multi edition bootable image       "
Print "    3) Inject drivers into one or more Windows editions and create a multi edition bootable image               "
Print "    4) Inject boot-critical drivers into one or more Windows editions and create a multi edition bootable image "
Color 0, 10
Print "    5) Make or update a bootable drive from one or more Windows ISO images and Windows PE / RE images           "
Print "    6) Create a bootable Windows ISO image that can include multiple editions                                   "
Print "    7) Create a bootable ISO image from Windows files in a folder                                               "
Print "    8) Reorganize the contents of a Windows ISO image                                                           "
Color 0, 3
Print "    9) Get WIM info - display basic info for each WIM in an ISO image                                           "
Print "   10) Modify the NAME and DESCRIPTION values for entries in a WIM file                                         "
Color 0, 6
Print "   11) Export drivers from this system                                                                          "
Print "   12) Expand drivers supplied in a .CAB file                                                                   "
Print "   13) Create a Virtual Disk (VHDX)                                                                             "
Print "   14) Create a VHD, deploy Windows to it, and add it to the boot menu to make a dual boot configuration        "
Print "   15) Create a generic ISO image and inject files and folders into it                                          "
Print "   16) Cleanup files and folders                                                                                "
Color 0, 8
Print "   17) Exit                                                                                                     "
Color 0, 13
Print "   18) Exit help and return to main menu                                                                        "
Locate 27, 0
Color 15
Input "   Please select the item you would like help with by entering its number (18 returns to the main menu): ", MenuSelection

Select Case MenuSelection
    Case 1
        GoTo GeneralHelp
    Case 2
        GoTo HelpInjectUpdates
    Case 3
        GoTo HelpInjectDrivers
    Case 4
        GoTo HelpInjectBCD
    Case 5
        GoTo HelpMakeMultiBootImage
    Case 6
        GoTo HelpMakeBootDisk2
    Case 7
        GoTo HelpCreateBootableISOFromFiles
    Case 8
        GoTo HelpChangeOrder
    Case 9
        GoTo HelpGetWimInfo
    Case 10
        GoTo HelpNameAndDescription
    Case 11
        GoTo HelpExportDrivers
    Case 12
        GoTo HelpExpandDrivers
    Case 13
        GoTo HelpCreateVHDX
    Case 14
        GoTo HelpAddVHDtoBootMenu
    Case 15
        GoTo HelpCreateISOImage
    Case 16
        GoTo HelpGetFolderToClean
    Case 17
        GoTo HelpExit
    Case 18
        ChDir ProgramStartDir$: GoTo BeginProgram
End Select

' We arrive here if the user makes an invalid selection from the main menu

Cls
Color 14, 4
Print
Print "You have made an invalid selection.";
Color 15
Print " You need to make a selection by entering a number from 1 to 18."
Pause
GoTo ProgramHelp

' Help Topic: Get general help on the use of this program

GeneralHelp:

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
Print " Program Help - General Help ";
Locate 9, 1
Color 15
Print "    1) System requirements       "
Print "    2) Terminology"
Print "    3) Responding to the program"
Print "    4) Hard disk vs removable media"
Print "    5) Reviewing log files"
Print "    6) How answer files are handled"
Print "    7) Auto shutdown and program pause"
Print "    8) Antivirus exclusions"
Print
Color 0, 13
Print "    9) Return to main help menu "
Locate 27, 0
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
Print "images that contain an INSTALL.WIM file in the \sources folder, not an INSTALL.ESD. There is one exception which is"
Print "addressed in the help sections related to those sections where it is applicable."
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
Print "21H1 and 21H2 in the same project."
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
Pause
GoTo GeneralHelp

' Help Topic: Get general help on the use of this program > Responding to the program

Responding:
Cls
Print "Responding to the Program"
Print "========================="
Print
Print "When the program asks for a path, you can enclose paths that have spaces in quotes if you wish, but this is not"
Print "necessary. The program will handle paths either way."
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
Print "IMPORTANT: Because the program maintains a keyboard buffer, you should NOT press keys at random while the program is"
Print "performing an operation and has focus. As soon as an operation is completed the keys you pressed will be processed."
Print
Print "However, the program now has a powerful script recording and playback tool. When you select the option to inject"
Print "Windows updates, drivers, or boot-critical drivers, you will be given the option to record or playback a script. The"
Print "scripts that you record will be saved to the folder in which the program is located as WIM_SCRIPT.TXT. If you wish to"
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
Print "IMPORTANT: Enter the index numbers from low to high. Don't specify a lower index number after a higher index number."
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
Print "The most important log file is the "; Chr$(34); "SANITIZED_ERROR_SUMMARY.log"; Chr$(34); " file. Another log file, the "; Chr$(34); "ERROR_SUMMARY.log"; Chr$(34); " file is"
Print "a summary of all errors generated by the use of the Microsoft DISM utility. At the time of this writing, Microsoft is"
Print "aware that there are issues causing errors to be generated. As a result, you are certain to see errors in this log."
Print
Print "The program now features a routine that will sanitize the log file and clean it up by removing errors that can be"
Print "ignored. The results are found in the file "; Chr$(34); "SANITIZED_ERROR_SUMMARY.log"; Chr$(34); ". Normally, this is the only log file that you"
Print "should need to monitor after any routine to update Windows editions has been run."
Print
Print "The original raw logs from DISM will have names such as dism.log_x64_1.txt. There will be one such log for each Windows"
Print "edition updated, with the number after the second underscore representing the index number of that image. x86 editions"
Print "of Windows will have x86 in the file name rather than x64. These files can be very large and you will not typically"
Print "need to look at these. If you find errors in the SANITIZED_ERROR_SUMMARY.log, the errors listed there will reference"
Print "the log file from which the error was taken and the timestamp of the error. You can then refer to that log file to"
Print "review the error in the context of operations that were taking place at the time."
Print
Print "In addition, you will find log files with names such as "; Chr$(34); "x64_1UpdateResults.txt"; Chr$(34); ". These log files will allow you to see"
Print "what updates have been applied to your WIM images. One such log will be present for each edition present. Note that the"
Print "number after the underscore will match the index number of that Windows edition in the final image."
Print
Print "Finally, the log named OSCDIMG.log will hold the results of the OSCDIMG utility used to create the final windows image."
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
Print "  1) For the routine to make or update a bootable drive from a Windows ISO image, we will ask the user if they want to"
Print "     exclude an answer file if one exists."
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
Print "automatically shut down the system when it is done running and you can pause the execution of the program to free up"
Print "resources to other programs."
Print
Print "To have the program perform an automatic shutdown, create a file on your desktop named "; Chr$(34); "Auto_Shutdown.txt"; Chr$(34); ". If a file by"
Print "that name exists, a shutdown will be performed when those routines have finished. To pause program execution, create a"
Print "file on your desktop named "; Chr$(34); "WIM_Pause.txt"; Chr$(34); "."
Print
Print "Note that for auto shutdown, you can change your mind at any time. The existence of this file will only cause a shutdown"
Print "when the routine finishes running so you are free to create that file at any time even while the program is running, or"
Print "you can delete / rename the file if you decide at some point that you do not want an automatic shutdown after all."
Print
Print "While the program is running, you will see a status indication on the upper right of the screen to remind you whether an"
Print "automatic shutdown will occur or not. If program execution is paused, a flashing message will be displayed as a reminder"
Print "that the program is paused. Note that when you make a change, the status will not update immediately. The status is"
Print "updated each time the program advances to the next item in the displayed checklist. However, when the program is resumed"
Print "by deleting or renaming the "; Chr$(34); "WIM_Pause.txt"; Chr$(34); " file, this status change will be reflected immediately."
Print
Print "Note that when the system is shut down automatically, any status messages or warning that would normally be displayed"
Print "will not be shown due to the shutdown. Instead, this information is logged to a file. The next time the program is run"
Print "that information will be automatically displayed. After viewing this information, that file will be automatically"
Print "deleted."
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
Print " Dual Architecture Edition         "
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
Locate 27, 0
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
Print "All Windows ISO images used need to have an INSTALL.WIM (not an INSTALL.ESD). The one exception to this is that the file"
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
Print "10 Version 21H1"; Chr$(34); " (or whatever your version is) and then click on the Last Updated column to sort with the latest"
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
Print "that is NOT described as a dynamic update."
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
Print "The advantage of this method is that a Boot Foundation ISO image is small,less than 50 MB. This means that the huge dual"
Print "architecture ISO image does not need to be kept."
Print
Print "   Download the dual architecture image as noted in option 1"
Print "   Copy all files and folders EXCEPT the x64 and x86 folder to a temporary location."
Print "   Create an ISO image from the folders and files in the temporary location (this program has a routine to create a"
Print "   generic ISO image that can be used for this). Make sure to put these files and folders at the root of the ISO image."
Print
Print "That ISO image is your Boot Foundation ISO image."
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
Print "update files, then create subfolders that look like this:"
Print
Print "\Answer_File        < Place autounattend.xml answer file here if you want to include one"
Print "\x64                < Folders with x64 (64-bit drivers will be organized under this folder"
Print "      \LCU          < Place the LCU (Latest Cumulative Update) in this folder"
Print "      \MicroCode    < Keep CPU microcode updates here"
Print "      \Other        < Other updates such as .NET update or Flash Player update"
Print "      \PE_Files     < Used to copy generic files (not Windows updates) to the Win PE image"
Print "      \SafeOS_DU    < Safe OS Dynamic Update - Used to update Windows Recovery (WinRE)"
Print "      \Setup_DU     < Setup Dynamic Update"
Print "\x86                < Folders with x86 (32-bit) drivers will be found here. Only used in Dual Architecture projects."
Print "      \LCU"
Print "      \MicroCode"
Print "      \Other"
Print "      \PE_Files"
Print "      \SafeOS_DU"
Print "      \Setup_DU"
Print
Print "Note that we used to have a folder for the SSU (Servicing Stack Update). Microsoft now combines this with the LCU."
Pause
Cls
Print "Description of Each Folder in the Structure"
Print "==========================================="
Print
Print "Answer_File - An autounattend.xml file will not be included from your original Windows image file(s). If you want to"
Print "include one in the final image, place it into the Answer_File folder."
Print
Print "LCU - Installs the ";: Color 0, 10: Print "L";: Color 15: Print "atest ";: Color 0, 10: Print "C";: Color 15: Print "umulative quality ";: Color 0, 10: Print "U";: Color 15: Print "pdate. The Title field in the update catalog will show Cumulative Update"
Print "for Windows 10 (or 11) for this update. Store only one LCU file in this folder. ";: Color 0, 10: Print "DO NOT";: Color 15: Print " download the LCU described as"
Print "being a Dynamic Update."
Print
Print "MicroCode - Unlike the other folders, the program will take no action on the contents of this folder. This folder is"
Print "merely a place to keep the latest MicroCode update file. We do this because the MicroCode updates do not apply to every"
Print "CPU so this is a good place to keep it if it will not be used. If you want the program to update your Windows editions"
Print "with the MicroCode update, simply copy the MicroCode update file from the \MicroCode folder to the \Other folder before"
Print "running the program."
Print
Print "Other - Updates that do not fall into the category of any other folder here are placed into the \Other folder. These"
Print "typically include .NET updates as an example. You can place multiple files in this folder, but you should save only the"
Print "latest version of each update type. For example, save only one .NET update here."
Print
Print "PE_Files - No Microsoft update files are placed into this folder. Instead, you will place generic files that need to be"
Print "accessible to Windows PE during Windows setup in this folder. This would typically include things like scripts. Simply"
Print "place any such files in this folder. If you later run Windows setup from media created these files, you will find the"
Print "files that you placed here on X:\ which is the RAM Drive that Windows setup creates during installation. To delete a"
Print "file from the WinPE image, create a dummy file with the same name as the file that you want to delete preceded with a"
Print "minus sign (-). For example, to delete a file called MyScript.bat, create any file in the \PE_Files folder, then rename"
Print "the file to -MyScript.bat. The case of the filename does not matter. You should rarely, if ever, need to use this."
Pause
Cls
Print "SafeOS_DU - Fixes for the Safe OS that are used to update Windows recovery environment (WinRE). Save your Safe OS"
Print "Dynamic Update file here. When downloading files from the Microsoft Update Catalog, the Safe OS Dynamic Update will"
Print "specifically indicate Safe OS Dynamic Update in the Title field."
Print
Print "Setup_DU - Fixes to Setup binaries or any files that Setup uses for feature updates. Note that when downloading files"
Print "from the Microsoft Update Catalog, the Setup Dynamic Update will indicate Windows 10 Dynamic Update in the Product field"
Print "and Dynamic Update for Windows 10 in the Title field. Store only one Setup Dynamic Update file in this folder."
Print
Print "SSU - The SSU or Servicing Stack Update contains fixes that are necessary to address the Windows 10 servicing stack and"
Print "is required to complete a feature update. These updates will show Servicing Stack Update in the Title field on the"
Print "update catalog. Only one SSU update should be stored in this folder."
Print
Print "TIP: If you have a new update that you wish the apply to your Windows edition(s) and you already have other updates"
Print "applied to your Windows edition(s), create an update folder with only the new update(s) and remove the contents of all"
Print "the other folders. This will cause the updates to be applied faster since the other updates don't have to be parsed to"
Print "see if their contents have already been applied. For example, if you download a new LCU (Latest Cumulative Update) and"
Print "you have already previously applied the other available updates such as the SSU, the Other updates category, etc., then"
Print "you can create a folder with only an LCU subfolder and the LCU placed in that folder. Note that this is not mandatory."
Print "It is perfectly fine to have updates in your update folders that are already applied, it will simply take longer."
Pause
GoTo HelpInjectUpdates

' Help Topic: Inject drivers into one or more Windows editions and create a multi edition bootable image

HelpInjectDrivers:

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
Locate 27, 0
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
Print "As with the routine to inject Windows updates, this routine requires Windows ISO images with an INSTALL.WIM and not an"
Print "INSTALL.ESD. Please see the help for the routine that injects Windows updates for a more detailed discussion of"
Print "acceptable Windows images."
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
Print " Dual Architecture Edition         "
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
Print "details."
Pause
GoTo ProgramHelp

' Help Topic: Make or update a bootable drive from one or more Windows ISO images and Windows PE / RE images
' or update an already existing drive

HelpMakeMultiBootImage:

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
Print " Program Help - Make a bootable drive from one or more Windows ";
Locate 4, 38
Print "                ISO images and Windows PE / RE images          ";
Locate 9, 1
Color 15
Print "    1) General information about this routine"
Print "    2) Disk limitations"
Print
Color 0, 13
Print "    3) Return to main help menu "
Locate 27, 0
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

' Help Topic: Make or update a bootable drive from one or more Windows ISO images and Windows PE / RE images
' or update an already existing drive > General information about this routine

MakeBootDriveHelp:

Cls
Print "General Information About This Routine"
Print "======================================"
Print
Print "This routine now has two major options - review the information related to the option that you wish to choose on the"
Print "pages below."
Pause
Cls
Print "Option 1 - MBR Boot Media"
Print "========================="
Print
Print "This routine will allow you to create bootable media from a bootable Windows ISO image. You will also be given the"
Print "choice to create additional partitions on that media that can be used to store other data. If you do choose to create"
Print "additional partitions, this routine can also BitLocker encrypt those partitions for you if you wish."
Print
Print "This routine uses a method to create the bootable media that should allow it to work on any system so long as that"
Print "system allows booting from the type of media (thumb drive, SD card, external HD, SSD, etc.) that you are using. It"
Print "should work with x86 and x64 systems, and systems that are BIOS based or UEFI based. For systems that will only boot"
Print "from external media that is formatted with FAT, this boot method will work around the limitation of a 4 GB maximum file"
Print "size. As a result, you do not need to break up large Windows image files into smaller pieces. If you are interested in"
Print "obtaining any technical details on how this works, please send a private message to me (hsehestedt) on TenForums.com."
Print
Print "The first time that you make a drive bootable using this routine, you should choose the option to WIPE the disk. This"
Print "will erase all data from the disk and properly prepare it. In the future, you can choose the REFRESH option which will"
Print "allow you to update the media from a new ISO image but it will leave all the data on other partitions alone. This is"
Print "perfect for larger media such as an external SSD because you can use that device as a Windows install / recovery disk,"
Print "but you can still use the remaining space on the drive for other things without needing to ever worry about erasing"
Print "that other data when you want to update the bootable portion of the disk. Note that when you choose to refresh a disk,"
Print "if media previously created with this routine is found, it will be refreshed automatically without you having to choose"
Print "what drive to update. If more than one such drive is found, then you will be asked to identify the disk to be updated."
Pause
Cls
Print "Option 2 - GPT Boot Image"
Print "========================="
Print
Print "This mode has some advantages but works only on UEFI / x64 based systems. Media created using this method wil not work"
Print "on BIOS / x86 based systems. The advantages of this method:"
Print
Print "1) Can boot multiple different operating systems or applications."
Print "2) Allows for disks greater than 2TB in size."
Print "3) Can support more than 4 primary partitions (we support 15 with this app, limit is actually 128)."
Print
Print "You can add multiple Windows images such as Windows 10 and 11 to the same disk. In addition, you can add Windows RE / PE"
Print "based media such as rescue and recovery disks for Macrium Reflect, etc. When booting from a disk made with procedure,"
Print "your system will display one instance of the boot device for each bootable partition. Note that this boot entry is based"
Print "upon the hardware disk device, and not what the contents are, so unfortunately it won't be able to display a description"
Print "of what each boot entry is. You should make note of the order in which you add options. Note that for each operating"
Print "system entry you add, we will need to create two physical partitions. Win PE / RE media and generic partitions only"
Print "create one partition each. Note that while an operating system (Windows 10 or 11) will create two partitions, only one"
Print "boot item is shown from the UEFI menu. As an example, if you create media that has Windows 10 and 11 (two operating"
Print "systems), two Win PE / RE based programs, and one generic partition, you will see 4 selectable lines on the UEFI boot"
Print "menu and seven partitions will be created."
Pause

GoTo HelpMakeMultiBootImage

' Help Topic: Make or update a bootable drive from one or more Windows ISO images and Windows PE / RE images
' or update an already existing drive > Disk limitations

DiskLimitationsHelp:

Cls
Print "Disk Limitations"
Print "================"
Print
Print "Be aware that for greatest compatibility, you should use media that is no larger than 2 TB in size for option 1 (MBR"
Print "Boot Image). If you use media that is larger than 2 TB in size, the program will give you the option to initialize"
Print "the media to 2 TB in size for the greatest compatibility, or to initialize the disk to its full capacity but sacrificing"
Print "the ability to be booted on BIOS based systems."
Print
Print "For option 2 (GPT Boot Image), you are not limited to a 2TB size. In addition, you can have up to 128 primary"
Print "partitions rather than just 4. This program supports 15 partitions. Be aware that disks created in this mode only work"
Print "with UEFI / x64 based systems."
Pause
GoTo HelpMakeMultiBootImage

' Help Topic: Create a bootable Windows ISO image that can include multiple editions

HelpMakeBootDisk2:

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
Print " Dual Architecture Edition         "
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
Print " Dual Architecture Edition         "
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

' Help Topic: Get WIM info - display basic info for each WIM in an ISO image

HelpGetWimInfo:

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
Print " Program Help - Get WIM info - display basic info for each WIM in an ISO image ";
Locate 9, 1
Color 15
Print "There are times where it may be necessary to know what index number is associated with a particular Windows edition,"
Print "or how many editions are stored in an image, or to view the NAME and DESCRIPTION metadata for Windows editions. This"
Print "routine will display that information and optionally save the output to a text file."
Pause
GoTo ProgramHelp

' Help Topic: Modify the NAME and DESCRIPTION values for entries in a WIM file

HelpNameAndDescription:

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
Print " Dual Architecture Edition         "
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
Print " Dual Architecture Edition         "
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
Print " Dual Architecture Edition         "
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
Pause
GoTo ProgramHelp

' Help Topic: Create a VHD, deploy Windows to it, and add it to the boot menu to make a dual boot configuration

HelpAddVHDtoBootMenu:

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
Print " Program Help - Create a VHD, deploy Windows to it, and add it to ";
Locate 4, 38
Print "                the boot menu to make a dual boot configuration   ";
Locate 9, 1
Color 15
Print "This routine can create an entirely new installation of Windows on a VHD that can be booted on a physical machine"
Print "without the need for any virtualization software. Since this copy of Windows runs from a VHD it requires no separate"
Print "disks or partitions on the system."
Pause
GoTo ProgramHelp

' Help Topic: Create a generic ISO image and inject files and folders into it

HelpCreateISOImage:

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
Print " Dual Architecture Edition         "
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

' Help Topic: Exit

HelpExit:

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
Print " Program Help - Exit ";
Locate 9, 1
Color 15
Print "Take a wild guess."
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
' Note that this routine creates a file called WIM_Info.txt. If you don't need it after
' a return from this subroutine, make sure to delete it.
'
' NOTE: Before calling the subroutine, set Silent$ to "Y" or "N". If set to "Y" then the routine will
' run silently. In other words, it save results to the WIM_Info.txt file but will not display it.

If Silent$ = "N" Then
    Cls
    Print "Preparing to display a list of indices...."
    Print
Else
    Print
    Print "Building a list of available editions."
End If

MountISO SourcePath$

Shell "echo File Name: " + Mid$(SourcePath$, _InStrRev(SourcePath$, "\") + 1) + " > WIM_Info.txt"
Shell "echo. >> WIM_Info.txt"
Shell "echo ***************************************************************************************************** >> WIM_Info.txt"
Shell "echo * Below is the list of Windows editions and associated indicies available for the above named file. * >> WIM_Info.txt"
Shell "echo ***************************************************************************************************** >> WIM_Info.txt"
Shell "echo. >> WIM_Info.txt"

If ArchitectureChoice$ <> "" Then
    Shell "echo. >> WIM_Info.txt"
    Shell "echo **************** >> WIM_Info.txt"
    Shell "echo * " + Right$(ArchitectureChoice$, 3) + " Editions * >> WIM_Info.txt"
    Shell "echo **************** >> WIM_Info.txt"
    Shell "echo. >> WIM_Info.txt"
End If

' The lines below test to see if this image has an install.esd or an install.wim and runs the appropriate command.
' Normally, we should not need this. Only an install.wim should be present for this project, but this routine can handle either.

InstallFileTest$ = MountedImageDriveLetter$ + ArchitectureChoice$ + "\sources\install.wim"

If _FileExists(InstallFileTest$) Then
    InstallFile$ = "\sources\install.wim >> WIM_Info.txt"
Else
    InstallFile$ = "\sources\install.esd >> WIM_Info.txt"
End If

Cmd$ = "dism /Get-WimInfo /WimFile:" + MountedImageDriveLetter$ + ArchitectureChoice$ + InstallFile$
Shell Cmd$

' Dismount the ISO image

Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourcePath$ + "'" + Chr$(34) + Chr$(34)
Shell _Hide Cmd$

' Display WIM_Info.txt which lists all of the indicies

If Silent$ = "N" Then
    Cls
    DisplayFile "WIM_Info.txt"
End If

Return


' *********************************************************************
' * The following subroutine creates registry files that are used to  *
' * recreate /boot/bcd and /efi/microsoft/boot/bcd files dynamically  *
' * when a dual architecture image is not available to retrieve those *
' * files.                                                            *
' *********************************************************************

Create_Reg_Files:

Temp$ = TempLocation$ + "\bcd_bios.reg"
bcd_ff = FreeFile
Open Temp$ For Output As #bcd_ff

Print #bcd_ff, "Windows Registry Editor Version 5.00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Description]"
Print #bcd_ff, ""; Chr$(34); "KeyName"; Chr$(34); "="; Chr$(34); "BCD00000001"; Chr$(34); ""
Print #bcd_ff, ""; Chr$(34); "GuidCache"; Chr$(34); "=hex:cc,b4,3c,7c,09,4c,d7,01,0f,27,00,00,d6,7a,a4,6b,b6,3f,c4,b9,24,\"
Print #bcd_ff, "  18,79,79"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20100000"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}\Elements\16000020]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{4636856e-540f-4170-a130-a84776f4c654}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20100000"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Elements\15000011]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Elements\15000013]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Elements\15000014]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,c2,01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20200003"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,37,00,65,00,61,00,32,00,65,00,31,00,61,00,63,00,2d,00,\"
Print #bcd_ff, "  32,00,65,00,36,00,31,00,2d,00,34,00,37,00,32,00,38,00,2d,00,61,00,61,00,61,\"
Print #bcd_ff, "  00,33,00,2d,00,38,00,39,00,36,00,64,00,39,00,64,00,30,00,61,00,39,00,66,00,\"
Print #bcd_ff, "  30,00,65,00,7d,00,00,00,7b,00,37,00,66,00,66,00,36,00,30,00,37,00,65,00,30,\"
Print #bcd_ff, "  00,2d,00,34,00,33,00,39,00,35,00,2d,00,31,00,31,00,64,00,62,00,2d,00,62,00,\"
Print #bcd_ff, "  30,00,64,00,65,00,2d,00,30,00,38,00,30,00,30,00,32,00,30,00,30,00,63,00,39,\"
Print #bcd_ff, "  00,61,00,36,00,36,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:30000000"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}\Elements\31000003]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,05,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}\Elements\32000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\boot\\boot.sdi"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:10200003"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\11000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:c8,dc,19,76,fe,fa,d9,11,b4,11,00,04,76,eb,a2,5f,00,00,00,00,01,\"
Print #bcd_ff, "  00,00,00,a8,00,00,00,00,00,00,00,03,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,01,00,00,00,80,00,00,00,05,00,00,00,05,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,5c,00,78,00,36,00,\"
Print #bcd_ff, "  34,00,5c,00,73,00,6f,00,75,00,72,00,63,00,65,00,73,00,5c,00,62,00,6f,00,6f,\"
Print #bcd_ff, "  00,74,00,2e,00,77,00,69,00,6d,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\12000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\windows\\system32\\boot\\winload.exe"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Windows 10 Setup (64-bit)"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\12000005]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "en-US"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,36,00,65,00,66,00,62,00,35,00,32,00,62,00,66,00,2d,00,\"
Print #bcd_ff, "  31,00,37,00,36,00,36,00,2d,00,34,00,31,00,64,00,62,00,2d,00,61,00,36,00,62,\"
Print #bcd_ff, "  00,33,00,2d,00,30,00,65,00,65,00,35,00,65,00,66,00,66,00,37,00,32,00,62,00,\"
Print #bcd_ff, "  64,00,37,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\21000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:c8,dc,19,76,fe,fa,d9,11,b4,11,00,04,76,eb,a2,5f,00,00,00,00,01,\"
Print #bcd_ff, "  00,00,00,a8,00,00,00,00,00,00,00,03,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,01,00,00,00,80,00,00,00,05,00,00,00,05,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,5c,00,78,00,36,00,\"
Print #bcd_ff, "  34,00,5c,00,73,00,6f,00,75,00,72,00,63,00,65,00,73,00,5c,00,62,00,6f,00,6f,\"
Print #bcd_ff, "  00,74,00,2e,00,77,00,69,00,6d,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\22000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\windows"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\250000c2]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\26000010]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\26000022]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\260000b0]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20100000"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,34,00,36,00,33,00,36,00,38,00,35,00,36,00,65,00,2d,00,\"
Print #bcd_ff, "  35,00,34,00,30,00,66,00,2d,00,34,00,31,00,37,00,30,00,2d,00,61,00,31,00,33,\"
Print #bcd_ff, "  00,30,00,2d,00,61,00,38,00,34,00,37,00,37,00,36,00,66,00,34,00,63,00,36,00,\"
Print #bcd_ff, "  35,00,34,00,7d,00,00,00,7b,00,30,00,63,00,65,00,34,00,39,00,39,00,31,00,62,\"
Print #bcd_ff, "  00,2d,00,65,00,36,00,62,00,33,00,2d,00,34,00,62,00,31,00,36,00,2d,00,62,00,\"
Print #bcd_ff, "  32,00,33,00,63,00,2d,00,35,00,65,00,30,00,64,00,39,00,32,00,35,00,30,00,65,\"
Print #bcd_ff, "  00,35,00,64,00,39,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20200003"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Hypervisor Settings"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements\250000f3]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements\250000f4]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements\250000f5]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,c2,01,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:10100002"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Windows Boot Manager"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\12000005]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "en-US"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,37,00,65,00,61,00,32,00,65,00,31,00,61,00,63,00,2d,00,\"
Print #bcd_ff, "  32,00,65,00,36,00,31,00,2d,00,34,00,37,00,32,00,38,00,2d,00,61,00,61,00,61,\"
Print #bcd_ff, "  00,33,00,2d,00,38,00,39,00,36,00,64,00,39,00,64,00,30,00,61,00,39,00,66,00,\"
Print #bcd_ff, "  30,00,65,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\23000003]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "{7619dcc9-fafe-11d9-b411-000476eba25f}"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\24000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,37,00,36,00,31,00,39,00,64,00,63,00,63,00,39,00,2d,00,\"
Print #bcd_ff, "  66,00,61,00,66,00,65,00,2d,00,31,00,31,00,64,00,39,00,2d,00,62,00,34,00,31,\"
Print #bcd_ff, "  00,31,00,2d,00,30,00,30,00,30,00,34,00,37,00,36,00,65,00,62,00,61,00,32,00,\"
Print #bcd_ff, "  35,00,66,00,7d,00,00,00,7b,00,62,00,61,00,37,00,66,00,34,00,64,00,62,00,63,\"
Print #bcd_ff, "  00,2d,00,62,00,37,00,66,00,63,00,2d,00,31,00,31,00,65,00,62,00,2d,00,62,00,\"
Print #bcd_ff, "  61,00,64,00,36,00,2d,00,61,00,34,00,36,00,62,00,62,00,36,00,33,00,66,00,63,\"
Print #bcd_ff, "  00,34,00,62,00,39,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\24000010]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,62,00,32,00,37,00,32,00,31,00,64,00,37,00,33,00,2d,00,\"
Print #bcd_ff, "  31,00,64,00,62,00,34,00,2d,00,34,00,63,00,36,00,32,00,2d,00,62,00,66,00,37,\"
Print #bcd_ff, "  00,38,00,2d,00,63,00,35,00,34,00,38,00,61,00,38,00,38,00,30,00,31,00,34,00,\"
Print #bcd_ff, "  32,00,64,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\25000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:1e"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:10200005"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\11000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,05,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\12000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\boot\\memtest.exe"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Windows Memory Diagnostic"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\12000005]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "en-US"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,37,00,65,00,61,00,32,00,65,00,31,00,61,00,63,00,2d,00,\"
Print #bcd_ff, "  32,00,65,00,36,00,31,00,2d,00,34,00,37,00,32,00,38,00,2d,00,61,00,61,00,61,\"
Print #bcd_ff, "  00,33,00,2d,00,38,00,39,00,36,00,64,00,39,00,64,00,30,00,61,00,39,00,66,00,\"
Print #bcd_ff, "  30,00,65,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:10200003"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\11000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:c8,dc,19,76,fe,fa,d9,11,b4,11,00,04,76,eb,a2,5f,00,00,00,00,01,\"
Print #bcd_ff, "  00,00,00,a8,00,00,00,00,00,00,00,03,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,01,00,00,00,80,00,00,00,05,00,00,00,05,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,5c,00,78,00,38,00,\"
Print #bcd_ff, "  36,00,5c,00,73,00,6f,00,75,00,72,00,63,00,65,00,73,00,5c,00,62,00,6f,00,6f,\"
Print #bcd_ff, "  00,74,00,2e,00,77,00,69,00,6d,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\12000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\windows\\system32\\boot\\winload.exe"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Windows 10 Setup (32-bit)"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\12000005]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "en-US"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,36,00,65,00,66,00,62,00,35,00,32,00,62,00,66,00,2d,00,\"
Print #bcd_ff, "  31,00,37,00,36,00,36,00,2d,00,34,00,31,00,64,00,62,00,2d,00,61,00,36,00,62,\"
Print #bcd_ff, "  00,33,00,2d,00,30,00,65,00,65,00,35,00,65,00,66,00,66,00,37,00,32,00,62,00,\"
Print #bcd_ff, "  64,00,37,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\21000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:c8,dc,19,76,fe,fa,d9,11,b4,11,00,04,76,eb,a2,5f,00,00,00,00,01,\"
Print #bcd_ff, "  00,00,00,a8,00,00,00,00,00,00,00,03,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,01,00,00,00,80,00,00,00,05,00,00,00,05,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,5c,00,78,00,38,00,\"
Print #bcd_ff, "  36,00,5c,00,73,00,6f,00,75,00,72,00,63,00,65,00,73,00,5c,00,62,00,6f,00,6f,\"
Print #bcd_ff, "  00,74,00,2e,00,77,00,69,00,6d,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\22000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\windows"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\250000c2]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\26000010]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\26000022]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\bcd\Objects\{ba7f4dbc-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\260000b0]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00"

Close #bcd_ff

Temp$ = TempLocation$ + "\bcd_efi.reg"
bcd_ff = FreeFile
Open Temp$ For Output As #bcd_ff

Print #bcd_ff, "Windows Registry Editor Version 5.00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Description]"
Print #bcd_ff, ""; Chr$(34); "KeyName"; Chr$(34); "="; Chr$(34); "BCD00000001"; Chr$(34); ""
Print #bcd_ff, ""; Chr$(34); "GuidCache"; Chr$(34); "=hex:09,ff,3e,7c,09,4c,d7,01,0f,27,00,00,d6,7a,a4,6b,b6,3f,c4,b9,24,\"
Print #bcd_ff, "  18,79,79"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20100000"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{0ce4991b-e6b3-4b16-b23c-5e0d9250e5d9}\Elements\16000020]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{4636856e-540f-4170-a130-a84776f4c654}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20100000"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Elements\15000011]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Elements\15000013]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{4636856e-540f-4170-a130-a84776f4c654}\Elements\15000014]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,c2,01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20200003"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{6efb52bf-1766-41db-a6b3-0ee5eff72bd7}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,37,00,65,00,61,00,32,00,65,00,31,00,61,00,63,00,2d,00,\"
Print #bcd_ff, "  32,00,65,00,36,00,31,00,2d,00,34,00,37,00,32,00,38,00,2d,00,61,00,61,00,61,\"
Print #bcd_ff, "  00,33,00,2d,00,38,00,39,00,36,00,64,00,39,00,64,00,30,00,61,00,39,00,66,00,\"
Print #bcd_ff, "  30,00,65,00,7d,00,00,00,7b,00,37,00,66,00,66,00,36,00,30,00,37,00,65,00,30,\"
Print #bcd_ff, "  00,2d,00,34,00,33,00,39,00,35,00,2d,00,31,00,31,00,64,00,62,00,2d,00,62,00,\"
Print #bcd_ff, "  30,00,64,00,65,00,2d,00,30,00,38,00,30,00,30,00,32,00,30,00,30,00,63,00,39,\"
Print #bcd_ff, "  00,61,00,36,00,36,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:30000000"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}\Elements\31000003]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,05,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc8-fafe-11d9-b411-000476eba25f}\Elements\32000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\boot\\boot.sdi"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:10200003"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\11000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:c8,dc,19,76,fe,fa,d9,11,b4,11,00,04,76,eb,a2,5f,00,00,00,00,01,\"
Print #bcd_ff, "  00,00,00,a8,00,00,00,00,00,00,00,03,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,01,00,00,00,80,00,00,00,05,00,00,00,05,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,5c,00,78,00,36,00,\"
Print #bcd_ff, "  34,00,5c,00,73,00,6f,00,75,00,72,00,63,00,65,00,73,00,5c,00,62,00,6f,00,6f,\"
Print #bcd_ff, "  00,74,00,2e,00,77,00,69,00,6d,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\12000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\windows\\system32\\boot\\winload.efi"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Windows 10 Setup (64-bit)"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\12000005]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "en-US"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,36,00,65,00,66,00,62,00,35,00,32,00,62,00,66,00,2d,00,\"
Print #bcd_ff, "  31,00,37,00,36,00,36,00,2d,00,34,00,31,00,64,00,62,00,2d,00,61,00,36,00,62,\"
Print #bcd_ff, "  00,33,00,2d,00,30,00,65,00,65,00,35,00,65,00,66,00,66,00,37,00,32,00,62,00,\"
Print #bcd_ff, "  64,00,37,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\16000060]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\21000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:c8,dc,19,76,fe,fa,d9,11,b4,11,00,04,76,eb,a2,5f,00,00,00,00,01,\"
Print #bcd_ff, "  00,00,00,a8,00,00,00,00,00,00,00,03,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,01,00,00,00,80,00,00,00,05,00,00,00,05,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,5c,00,78,00,36,00,\"
Print #bcd_ff, "  34,00,5c,00,73,00,6f,00,75,00,72,00,63,00,65,00,73,00,5c,00,62,00,6f,00,6f,\"
Print #bcd_ff, "  00,74,00,2e,00,77,00,69,00,6d,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\22000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\windows"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\250000c2]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\26000010]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\26000022]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7619dcc9-fafe-11d9-b411-000476eba25f}\Elements\260000b0]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20100000"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ea2e1ac-2e61-4728-aaa3-896d9d0a9f0e}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,34,00,36,00,33,00,36,00,38,00,35,00,36,00,65,00,2d,00,\"
Print #bcd_ff, "  35,00,34,00,30,00,66,00,2d,00,34,00,31,00,37,00,30,00,2d,00,61,00,31,00,33,\"
Print #bcd_ff, "  00,30,00,2d,00,61,00,38,00,34,00,37,00,37,00,36,00,66,00,34,00,63,00,36,00,\"
Print #bcd_ff, "  35,00,34,00,7d,00,00,00,7b,00,30,00,63,00,65,00,34,00,39,00,39,00,31,00,62,\"
Print #bcd_ff, "  00,2d,00,65,00,36,00,62,00,33,00,2d,00,34,00,62,00,31,00,36,00,2d,00,62,00,\"
Print #bcd_ff, "  32,00,33,00,63,00,2d,00,35,00,65,00,30,00,64,00,39,00,32,00,35,00,30,00,65,\"
Print #bcd_ff, "  00,35,00,64,00,39,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:20200003"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Hypervisor Settings"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements\250000f3]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements\250000f4]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{7ff607e0-4395-11db-b0de-0800200c9a66}\Elements\250000f5]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,c2,01,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:10100002"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Windows Boot Manager"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\12000005]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "en-US"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,37,00,65,00,61,00,32,00,65,00,31,00,61,00,63,00,2d,00,\"
Print #bcd_ff, "  32,00,65,00,36,00,31,00,2d,00,34,00,37,00,32,00,38,00,2d,00,61,00,61,00,61,\"
Print #bcd_ff, "  00,33,00,2d,00,38,00,39,00,36,00,64,00,39,00,64,00,30,00,61,00,39,00,66,00,\"
Print #bcd_ff, "  30,00,65,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\23000003]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "{7619dcc9-fafe-11d9-b411-000476eba25f}"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\24000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,37,00,36,00,31,00,39,00,64,00,63,00,63,00,39,00,2d,00,\"
Print #bcd_ff, "  66,00,61,00,66,00,65,00,2d,00,31,00,31,00,64,00,39,00,2d,00,62,00,34,00,31,\"
Print #bcd_ff, "  00,31,00,2d,00,30,00,30,00,30,00,34,00,37,00,36,00,65,00,62,00,61,00,32,00,\"
Print #bcd_ff, "  35,00,66,00,7d,00,00,00,7b,00,62,00,61,00,38,00,31,00,39,00,37,00,66,00,39,\"
Print #bcd_ff, "  00,2d,00,62,00,37,00,66,00,63,00,2d,00,31,00,31,00,65,00,62,00,2d,00,62,00,\"
Print #bcd_ff, "  61,00,64,00,36,00,2d,00,61,00,34,00,36,00,62,00,62,00,36,00,33,00,66,00,63,\"
Print #bcd_ff, "  00,34,00,62,00,39,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\24000010]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,62,00,32,00,37,00,32,00,31,00,64,00,37,00,33,00,2d,00,\"
Print #bcd_ff, "  31,00,64,00,62,00,34,00,2d,00,34,00,63,00,36,00,32,00,2d,00,62,00,66,00,37,\"
Print #bcd_ff, "  00,38,00,2d,00,63,00,35,00,34,00,38,00,61,00,38,00,38,00,30,00,31,00,34,00,\"
Print #bcd_ff, "  32,00,64,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{9dea862c-5cdd-4e70-acc1-f32b344d4795}\Elements\25000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:1e"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:10200005"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\11000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,05,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\12000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\efi\\microsoft\\boot\\memtest.efi"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Windows Memory Diagnostic"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\12000005]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "en-US"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{b2721d73-1db4-4c62-bf78-c548a880142d}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,37,00,65,00,61,00,32,00,65,00,31,00,61,00,63,00,2d,00,\"
Print #bcd_ff, "  32,00,65,00,36,00,31,00,2d,00,34,00,37,00,32,00,38,00,2d,00,61,00,61,00,61,\"
Print #bcd_ff, "  00,33,00,2d,00,38,00,39,00,36,00,64,00,39,00,64,00,30,00,61,00,39,00,66,00,\"
Print #bcd_ff, "  30,00,65,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Description]"
Print #bcd_ff, ""; Chr$(34); "Type"; Chr$(34); "=dword:10200003"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements]"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\11000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:c8,dc,19,76,fe,fa,d9,11,b4,11,00,04,76,eb,a2,5f,00,00,00,00,01,\"
Print #bcd_ff, "  00,00,00,a8,00,00,00,00,00,00,00,03,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,01,00,00,00,80,00,00,00,05,00,00,00,05,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,5c,00,78,00,38,00,\"
Print #bcd_ff, "  36,00,5c,00,73,00,6f,00,75,00,72,00,63,00,65,00,73,00,5c,00,62,00,6f,00,6f,\"
Print #bcd_ff, "  00,74,00,2e,00,77,00,69,00,6d,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\12000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\windows\\system32\\boot\\winload.efi"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\12000004]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "Windows 10 Setup (32-bit)"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\12000005]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "en-US"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\14000006]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex(7):7b,00,36,00,65,00,66,00,62,00,35,00,32,00,62,00,66,00,2d,00,\"
Print #bcd_ff, "  31,00,37,00,36,00,36,00,2d,00,34,00,31,00,64,00,62,00,2d,00,61,00,36,00,62,\"
Print #bcd_ff, "  00,33,00,2d,00,30,00,65,00,65,00,35,00,65,00,66,00,66,00,37,00,32,00,62,00,\"
Print #bcd_ff, "  64,00,37,00,7d,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\16000060]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\21000001]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:c8,dc,19,76,fe,fa,d9,11,b4,11,00,04,76,eb,a2,5f,00,00,00,00,01,\"
Print #bcd_ff, "  00,00,00,a8,00,00,00,00,00,00,00,03,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,01,00,00,00,80,00,00,00,05,00,00,00,05,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,48,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\"
Print #bcd_ff, "  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,5c,00,78,00,38,00,\"
Print #bcd_ff, "  36,00,5c,00,73,00,6f,00,75,00,72,00,63,00,65,00,73,00,5c,00,62,00,6f,00,6f,\"
Print #bcd_ff, "  00,74,00,2e,00,77,00,69,00,6d,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\22000002]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "="; Chr$(34); "\\windows"; Chr$(34); ""
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\250000c2]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00,00,00,00,00,00,00,00"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\26000010]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\26000022]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:01"
Print #bcd_ff, ""
Print #bcd_ff, "[HKEY_LOCAL_MACHINE\BCD\Objects\{ba8197f9-b7fc-11eb-bad6-a46bb63fc4b9}\Elements\260000b0]"
Print #bcd_ff, ""; Chr$(34); "Element"; Chr$(34); "=hex:00"

Close #bcd_ff

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


Sub CleanPath (Path$)

    ' Remove trailing backslash from a path

    ' To use this subroutine: Pass the path to this sub, the sub will return the path
    ' without a trailing backslash in Temp$.

    Temp$ = Path$

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

    Dim Cmd As String
    Dim file As String

    NumberOfFiles = 0 ' Set initial value
    FileType$ = UCase$(FileType$)

    ' Build the command to be run

    If FileType$ = "*" Then
        Select Case SearchSubFolders$
            Case "N"
                Cmd$ = "DIR /B " + Chr$(34) + Path$ + Chr$(34) + " > WIM_TEMP.TXT"
            Case "Y"
                Cmd$ = "DIR /B " + Chr$(34) + Path$ + Chr$(34) + " /s" + " > WIM_TEMP.TXT"
        End Select
    Else
        Select Case SearchSubFolders$
            Case "N"
                Cmd$ = "DIR /B " + Chr$(34) + Path$ + "*" + FileType$ + Chr$(34) + " > WIM_TEMP.TXT"
            Case "Y"
                Cmd$ = "DIR /B " + Chr$(34) + Path$ + "*" + FileType$ + Chr$(34) + " /s" + " > WIM_TEMP.TXT"
        End Select
    End If

    Shell _Hide Cmd$

    If _FileExists("WIM_TEMP.TXT") Then
        Open "WIM_TEMP.TXT" For Input As #1
        Do Until EOF(1)
            Line Input #1, file$
            If FileType$ = "*" Then
                If file$ <> "File Not Found" Then
                    NumberOfFiles = NumberOfFiles + 1

                    If Left$(file$, 1) = "-" Then
                        TempArray$(NumberOfFiles) = "-" + Path$ + Right$(file$, (Len(file$) - 1))
                    Else
                        TempArray$(NumberOfFiles) = Path$ + file$
                    End If
                End If
            ElseIf UCase$(Right$(file$, 4)) = UCase$(FileType$) Then
                NumberOfFiles = NumberOfFiles + 1

                ' In case we are injecting drivers, we would be searching for ".INF" files here. For these files, we have no reason to store the name of these files
                ' because we don't process the files one by one. All we need for these is confirmation that .INF files exist.

                If FileType$ <> ".INF" Then
                    TempArray$(NumberOfFiles) = Path$ + file$
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

    Dim Cmd As String
    Dim GetLine As String
    Dim count As Integer

    MountedImageCDROMID$ = ""
    MountedImageDriveLetter$ = ""

    CleanPath (ImagePath$)
    ImagePath$ = Temp$
    Cmd$ = "powershell.exe -command " + Chr$(34) + "Mount-DiskImage " + Chr$(34) + "'" + ImagePath$ + "'" + Chr$(34) + Chr$(34) + " > MountInfo1.txt"
    Shell Cmd$
    Cmd$ = "powershell.exe -command " + Chr$(34) + "Get-DiskImage -ImagePath '" + ImagePath$ + "' | Get-Volume" + Chr$(34) + " > MountInfo2.txt"
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

    Dim ChosenIndexString As String
    Dim Cmd As String
    Dim ReadLine As String
    Dim position As Integer

    ChosenIndexString$ = Str$(ChosenIndex)
    ChosenIndexString$ = Right$(ChosenIndexString$, ((Len(ChosenIndexString$) - 1)))

    ' Clear variable

    MountedImageDriveLetter$ = ""
    MountISO SourcePath$

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

    Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourcePath$ + "'" + Chr$(34) + Chr$(34)
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

    ' This routine will save WIM info to a text file called WIM_Info.txt located in the same
    ' folder that the program is run from. This routine can be run with status information or
    ' silently. Pass to this routine the name of the ISO image for which to get info in a
    ' string and a 0 or 1 to indicate silent mode (0 = run normally, 1 = run silent).
    '
    ' If you no longer need the file WIM_Info.txt, make sure to delete it after a return from
    ' this routine.

    ' Declare local variables

    Dim Architecture As String ' Tracks the architecture type of the selected ISO image
    Dim Cmd As String ' Holds a string that has been built to be run with a "Shell" command
    Dim InstallFile As String
    Dim InstallFileTest As String
    Dim LocalTemp As String ' Temporary data

    Shell "echo File Name: " + Mid$(SourcePath$, _InStrRev(SourcePath$, "\") + 1) + " > WIM_Info.txt"
    Shell "echo. >> WIM_Info.txt"
    Shell "echo ***************************************************************************************************** >> WIM_Info.txt"
    Shell "echo * Below is the list of Windows editions and associated indicies available for the above named file. * >> WIM_Info.txt"
    Shell "echo *                                                                                                   * >> WIM_Info.txt"
    Shell "echo * If you are viewing this file on screen via the app and the info is more than one screen long,     * >> WIM_Info.txt"
    Shell "echo * press the SPACEBAR to advance one screen at a time, or ENTER to advance one line at a time.       * >> WIM_Info.txt"
    Shell "echo ***************************************************************************************************** >> WIM_Info.txt"
    Shell "echo. >> WIM_Info.txt"

    ' Determine if the file specified holds a dual architecture installation
    ' Unlike the routine for creating a multiboot disk, in this case we only
    ' need to know if the architecture is dual or single so the only values
    ' we use for ProjectArchitecture$ are DUAL or SINGLE.

    If GetWimInfo_Silent = 0 Then
        Cls
        Print "*************************************************************************"
        Print "* Checking to see if the image specified is dual or single architecture *"
        Print "*************************************************************************"
        Print
        Print "**************************"
        Print "* Mounting the ISO image *"
        Print "**************************"
        Print
    End If

    MountISO SourcePath$

    LocalTemp$ = MountedImageDriveLetter$ + "\x64"

    If _DirExists(LocalTemp$) Then
        Architecture$ = "DUAL"
    Else
        Architecture$ = "SINGLE"
    End If

    If Architecture$ = "SINGLE" Then
        Shell "echo *************************************** >> WIM_Info.txt"
        Shell "echo * This is a single architecture image * >> WIM_Info.txt"
        Shell "echo *************************************** >> WIM_Info.txt"
        Shell "echo. >> WIM_Info.txt"
        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Get-WimInfo /WimFile:" + Chr$(34) + MountedImageDriveLetter$ + "\sources\install.wim" + Chr$(34) + " >> WIM_Info.txt"
        Shell Chr$(34) + Cmd$ + Chr$(34)
    Else
        Shell "echo ************************************* >> WIM_Info.txt"
        Shell "echo * This is a dual architecture image * >> WIM_Info.txt"
        Shell "echo ************************************* >> WIM_Info.txt"
        Shell "echo. >> WIM_Info.txt"
        Shell "echo     **************** >> WIM_Info.txt"
        Shell "echo     * x86 Editions * >> WIM_Info.txt"
        Shell "echo     **************** >> WIM_Info.txt"
        Shell "echo. >> WIM_Info.txt"

        ' The lines below test to see if this dual architecture image has an install.esd or an install.wim and runs the appropriate command.
        ' Normally, we should not need this. Only an install.wim should be present for this project, but this routine can handle either.

        InstallFileTest$ = MountedImageDriveLetter$ + "\x86\sources\install.wim"
        If _FileExists(InstallFileTest$) Then
            InstallFile$ = "\x86\sources\install.wim"
        Else
            InstallFile$ = "\x86\sources\install.esd"
        End If

        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Get-WimInfo /WimFile:" + Chr$(34) + MountedImageDriveLetter$ + InstallFile$ + Chr$(34) + " >> WIM_Info.txt"
        Shell Chr$(34) + Cmd$ + Chr$(34)
        Shell "echo. >> WIM_Info.txt"
        Shell "echo     **************** >> WIM_Info.txt"
        Shell "echo     * x64 Editions * >> WIM_Info.txt"
        Shell "echo     **************** >> WIM_Info.txt"
        Shell "echo. >> WIM_Info.txt"

        ' The lines below test to see if this dual architecture image has an install.esd or an install.wim and runs the appropriate command.
        ' Normally, we should not need this. Only an install.wim should be present for this project, but this routine can handle either.

        If _FileExists(InstallFileTest$) Then
            InstallFile$ = "\x64\sources\install.wim"
        Else
            InstallFile$ = "\x64\sources\install.esd"
        End If
        Cmd$ = Chr$(34) + DISMLocation$ + Chr$(34) + " /Get-WimInfo  /WimFile:" + Chr$(34) + MountedImageDriveLetter$ + InstallFile$ + Chr$(34) + " >> WIM_Info.txt"
        Shell Chr$(34) + Cmd$ + Chr$(34)
    End If

    If GetWimInfo_Silent = 0 Then
        Print "*************************"
        Print "* Dismounting the image *"
        Print "*************************"
        Print
    End If

    Cmd$ = "powershell.exe -command " + Chr$(34) + "Dismount-DiskImage " + Chr$(34) + "'" + SourcePath$ + "'" + Chr$(34) + Chr$(34) + " > NUL"
    Shell Cmd$

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

    ' Option 1

    ' The section of code below will display a message if the program detects pending installs in a Windows image. This section of code
    ' needs to be modified if either "Option 2" or "Option 3" below are enabled because they will conflict with each other as they are
    ' currently coded.

    Print OverallStatus$;
    If OpsPending$ = "Y" Then
        Print "   ";: Color 4: Print "PENDING INSTALLS DETECTED!";: Color 10: Print " More info will be displayed when routine is done.": Color 15
    Else
        Print
    End If
    For x = 1 To Len(OverallStatus$)
        Print "*";
    Next x

    ' The 3 lines below that are commented out were displaying text on the screen illogically. I've commented them out for now but
    ' it should be okay to delete them permanently.

    '    If ErrorsWereFound$ = "Y" Then
    '        Print "                    ";: Color 10: Print "You can also review the ERROR_SUMMARY.log file now.": Color 15
    '    End If

    ' Option 2

    ' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    ' NOTE: The section of code below was used to notify a user if errors were detected even prior to the completion of this routine. After each
    ' edition of Windows was updated, we parse the DISM logs and look for errors. If errors were found, we would display this on the screen so that
    ' the user could already begin looking at some log files because it could take a very long time for all editions to be processed if the user
    ' queues up a lot of them. However, at the time of this writing, there is a problem with errors being logged that are not critical. This was
    ' not the case previously. I have reported this situation to Microsoft, they are aware of it, and "may" fix it in the future. This was
    ' several months ago already. Because of this problem, I am temporarily taking those message out. Please note that there is another message
    ' at the end of this routine that we will leave in place. In addition, we have added code to try to filter out most of the messages that can
    ' be ignored.

    ' To re-enable error display, uncomment the lines below between the "**********" lines and then follow the steps in next comment.
    ' **********

    'PRINT OverallStatus$;
    'IF ErrorsWereFound$ = "Y" THEN
    '    PRINT "   ";: COLOR 4: PRINT "ERRORS DETECTED!";: COLOR 10: PRINT " More info will be displayed when routine is done.": COLOR 15
    'ELSE
    '    PRINT
    'END IF
    'FOR x = 1 TO LEN(OverallStatus$)
    '    PRINT "*";
    'NEXT x
    'IF ErrorsWereFound$ = "Y" THEN
    '    PRINT "                    ";: COLOR 10: PRINT "You can also review the ERROR_SUMMARY.log file now.": COLOR 15
    'END IF

    ' **********

    ' End Option 2

    'Option 3

    ' To re-enable error display, comment out or remove the lines below between the "----------" lines.
    ' ----------

    'PRINT OverallStatus$
    'FOR x = 1 TO LEN(OverallStatus$)
    '    PRINT "*";
    'NEXT x
    'PRINT ""

    ' ----------

    ' End Option 3

    ' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


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
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[             ] Updating WinRE"
            Print "[             ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 2
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Mounting a Windows Edition": Color 15
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[             ] Updating WinRE"
            Print "[             ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 3
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Updating WinRE"
            Print "[             ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 4
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items": Color 15
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Updating WinPE (Index 1 of 2)"
            Print "[             ] Updating WinPE (Index 2 of 2)"
        Case 5
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items": Color 15
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Updating WinPE (Index 2 of 2)"
        Case 6
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 7
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 8
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 9
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 10
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 11
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 12
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 13
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 14
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Color 0, 10: Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image": Color 15
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[  COMPLETED  ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
        Case 15
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Print "- Update Tasks: These tasks are performed for each Windows edition"
            Print "[  COMPLETED  ] Mounting a Windows Edition"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[  COMPLETED  ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[  COMPLETED  ] Locking in Updates"
            Print "[  COMPLETED  ] Adding Other Updates and Setup Dynamic Updates"
            Print "[  COMPLETED  ] Creating Log Files"
            Print "[  COMPLETED  ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[  COMPLETED  ] Creating Base Image"
            Print "[  COMPLETED  ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[  COMPLETED  ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
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
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Moving Updated WIM Files to Base Image"
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
            Print "[  COMPLETED  ] Moving Updated WIM Files to Base Image"
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
            Print "[  COMPLETED  ] Moving Updated WIM Files to Base Image"
            Print "[  COMPLETED  ] Creating Final ISO Image"
        Case 25
            Print "- Pre-Update Task: This task is performed once before applying updates"
            Print "[  COMPLETED  ] Exporting All Windows Editions"
            Print
            Color 0, 10: Print "- Update Tasks: These tasks are performed for each Windows edition": Color 15
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Mounting a Windows Edition": Color 15
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 1"
            Print "[             ] Adding Servicing Stack and Cumulative Updates, Pass 2"
            Print "[             ] Locking in Updates"
            Print "[             ] Adding Other Updates and Setup Dynamic Updates"
            Print "[             ] Creating Log Files"
            Print "[             ] Unmounting and Saving Windows Edition"
            Print
            Print "- Post-Update Tasks: These tasks are performed after updating images to create the final ISO image"
            Print "[             ] Creating Base Image"
            Print "[             ] Moving Updated WIM Files to Base Image and Syncing File Versions"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items": Color 15
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
            Print "[             ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items": Color 15
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
            Print "[             ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Color 0, 10: Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items": Color 15
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
            Print "[             ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[             ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Moving Updated WIM Files to Base Image"
            Print "[             ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[  COMPLETED  ] Moving Updated WIM Files to Base Image"
            Print "[ ";: Color 10: Print "IN PROGRESS";: Color 15: Print " ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
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
            Print "[  COMPLETED  ] Moving Updated WIM Files to Base Image"
            Print "[  COMPLETED  ] Creating Final ISO Image"
            Print
            Print "- WinRE and WinPE: Updates applied upon finding first x64 and first x86 instance of these items"
            Print "[  COMPLETED  ] Updating WinRE"
            Print "[  COMPLETED  ] Updating WinPE (Index 1 of 2)"
            Print "[  COMPLETED  ] Updating WinPE (Index 2 of 2)"
    End Select

    ' Display the Auto Shutdown status. If a file named "Auto_Shutdown.txt" exists on the desktop, then the system will be
    ' shutdown when the program is done running. Note that this file can be created or removed / renamed by the user even
    ' while the program is running. The status will be updated in the display each time this status display routine is
    ' updated, which ocurrs with the start of each new step in the update process.
    '
    ' Also, check for the existance of a file named "WIM_PAUSE.txt" on the desktop. So long as that file exists, pause the
    ' execution of the program.

    Locate 1, 95: Print "Auto Shutdown: ";

    If _FileExists(Environ$("userprofile") + "\Desktop\Auto_Shutdown.txt") Then
        Color 14, 4: Print "Enabled";: Color 15
    Else
        Color 10: Print "Disabled";: Color 15
    End If

    Locate 29, 1: Color 0, 10: Print "   Place a file named AUTO_SHUTDOWN.TXT on desktop to shutdown system when done or WIM_PAUSE.TXT to pause the program   ";
    Locate 30, 1: Print "   Changes are reflected when progress advances to the next step                                                        ";: Color 15

    ' If the file "WIM_PAUSE.txt" exists on the desktop, pause program execution until
    ' that file is deleted or renamed.

    Do While _FileExists(Environ$("userprofile") + "\Desktop\WIM_PAUSE.txt")
        Locate 2, 95: Color 14, 4: Print "PROGRAM EXECUTION PAUSED";: Color 15
        _Delay .5
        Locate 2, 95: Print "PROGRAM EXECUTION PAUSED";
        _Delay .5
    Loop

    Locate 2, 95: Print "                        ";

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


Sub FindDISMLogErrors_SingleFile (Path$, LogPath$)


    ' Pass to this routine the full path including a file name of a DISM log file and it will scan the file for errors.
    ' Pass the location of the log files as the second parameter.
    ' If an error is found it will note this and will pass back to the main program a value of "Y" in DISM_Error_Found$.
    ' It will also create a log file called ERROR_SUMMARY.log
    ' NOTE: This routine does not check for the existance of the file being passed to it. You should check the validity of
    ' the path and file before calling this routine.

    DISM_Error_Found$ = "N"
    Dim ff As Long ' Hold the next open Free File number
    Dim ff2 As Long ' Used to get a file number for the 2nd file that needs to be open at the same time
    Dim Position1 As Double
    Position1 = 0
    Dim Position2 As Double
    Dim StartOfError As Double
    Dim LogFile As String
    Dim ErrorMessage As String

    ' Init variables

    ff = FreeFile
    Open Path$ For Binary As #ff
    LogFile$ = Space$(LOF(ff))
    Get #ff, 1, LogFile$
    Close #ff

    Do
        Position1 = InStr(Position1 + 1, LogFile$, "Error                 ")
        If Position1 Then
            StartOfError = Position1 - 21
            Position2 = InStr(StartOfError, LogFile$, (Chr$(13) + Chr$(10)))
            ErrorMessage$ = Mid$(LogFile$, StartOfError, (Position2 - StartOfError))
            DISM_Error_Found$ = "Y"
            ff2 = FreeFile
            Open (LogPath$ + "\ERROR_SUMMARY.log") For Append As #ff2
            Print #ff2, "Warning! Error was reported in the log file named:"
            Print #ff2, Path$
            Print #ff2, "The error reported is:"
            Print #ff2, ErrorMessage$
            Print #ff2, ""
            Close #ff2
        End If
    Loop Until Position1 = 0
End Sub


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
    Open "WIM_Info.txt" For Input As #1
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

    EiCfg:

    Cls
    Print "Type ";: Color 0, 10: Print "HELP";: Color 15: Print " if you need information about the below option."
    Print
    Print "Do you want to inject an ";: Color 0, 10: Print "EI.CFG";: Color 15: Print " file into your final image";: Input CreateEiCfg$

    If ScriptingChoice$ = "R" Then
        Print #5, ":: Do you want to inject an EI.CFG file into your final image?"
        If UCase$(Left$(CreateEiCfg$, 1)) = "H" Then
            Print #5, ":: Help for this option was requested."
            Print #5, "HELP"
        ElseIf CreateEiCfg$ = "" Then
            Print #5, "<ENTER>"
        Else
            Print #5, CreateEiCfg$
        End If
        Print #5, ""
    End If

    If UCase$(Left$(CreateEiCfg$, 1)) = "H" Then
        Cls
        Print "If you have multiple editions of Windows in an image, Windows setup may not ask you which edition to install. If your"
        Print "BIOS / firmware uses a signature to indicate the edition that originally shipped with the system it may simply force"
        Print "installation of that Windows edition. As axample, assume that you have a laptop that shipped with Windows 10 Home"
        Print "edition preinstalled. You upgrade the system to Windows 10 Professional. Eventually you decide that that you want to"
        Print "perform a clean install of Windows 10, or maybe even Windows 11. When you begin the installation you are given no"
        Print "choice of what Windows edition to install. Widows setup simply proceeds to install the Home edition of Windows because"
        Print "that is what the BIOS signature indicates was installed from the factory."
        Print
        Print "Injecting the EI.CFG file into your image will force setup to allow you to choose the edition of Windows to be"
        Print "installed if your image contains multiple editions."
        Print
        Print "Note that if you use an answer file to perform an unattended setup, this file will have no effect since the answer file"
        Print "specifies the edition to be installed."
        Pause
        GoTo EiCfg
    End If

    YesOrNo CreateEiCfg$
    CreateEiCfg$ = YN$

    If CreateEiCfg$ = "X" Then
        Print
        Color 14, 4
        Print "Please provide a valid response."
        Color 15

        If ScriptingChoice$ = "R" Then
            Print #5, ":: The above response was not valid."
            Print #5, ""
        End If

        Pause
        GoTo EiCfg
    End If

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
    ' NumberOfDisks - This will indicate the number of disks seen by the system
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
    ff = FreeFile
    Open "DiskpartOut2.txt" For Binary As #ff
    ListOfDisks$ = Space$(LOF(ff))
    Get #ff, 1, ListOfDisks$
    Close #ff
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
            Input "Script File Name: ", ScriptFile$
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




' Release Notes
'
' 7.5.2.28
' Moved release notes to end of program.
'
' Updated the "Make a bootable drive from a Windows ISO image or update an already existing drive" routine to display a note stating that a disk
' created with this routine will only use 2TB of capacity since the disk will be initialized as MBR rather than GPT. This is needed to make the
' disk bootable on both BIOS and UEFI based systems.
'
' Added ability for user to override this behavior so that the program will create the disk with GPT rather than MBR. This will allow a disk that
' is larger than 2TB to use its entire capacity but sacrifices the ability to boot on a BIOS based system.
'
' Compiled with QB64 Feb 10, 2020 dev build
'
' 7.5.2.29
' Added a check at the end of the "Make a bootable drive from a Windows ISO image or update an already existing drive" routine to parse any DISM.LOG files and
' check for potentional errors. To do this, we read each line in the log and check for the word "Error" starting at the 22nd character in the line.
'
' Compiled with QB64 Feb 10, 2020 dev build
'
' 7.6.0.30 - Feb 13, 2020
' After testing the changes made in 7.5.2.29 I noticed that DISM log files were being overwritten. For example, if you update 5 editions of Windows, 5 seperate
' log files are created but only a DISM.LOG and DISM.LOG.BAK are kept. I have modified the code to rename each log file after it is created so that all log files
' are maintained, at least until the program is run again at which time it will purge old logs. As a result of now maintaining all these logs, it can take a very
' long time to parse all the log files when checking for errors. Because of this, I give the user an option to skip the log file check for errors, although I do
' suggest to them that they allow the check to take place.
'
' I also discovered that Console Title Bar messages were not getting properly displayed. This is because one of the items that I was displaying was the Program
' Version which I store in ProgramVersion$. I had to DIM this variable as shared so that it would be accessible in SUB sections of code.
'
' Compiled on the new QB64 version 1.4, released on the same day as this build.
'
' 7.6.0.31 - Feb 13, 2020
' Since the log files scan at the end of the "Make a bootable drive from a Windows ISO image or update an already existing drive" routine can take a long time to
' process, I have updated the routine to terminate as soon as an error is detected. That way, if we find an error fairly early on, we can save a lot of time since
' there is no reason to read the rest of the current log or any further log files. In addition, we now report what log file the error was found in and the line
' number within that log where the error was detected along with the contents of that line.
'
' Compiled on QB64 Ver. 1.4
'
' 7.6.1.32 - Feb 14, 2020
' At the start and end of the DISM log file scan for errors for the "Make a bootable drive from a Windows ISO image or update an already existing drive" routine,
' Display the date and time that the operation was started and completed. I may pull this code eventually, but for now it will serve as a way to gauge how long
' this process takes. That way, if I tweak the error scanning code I can better tell if this is improving or hurting performance.
'
' Compiled on QB64 Ver. 1.4
'
' 8.0.0.33 - Feb 15, 2020
' This is a major update. The sections called  "Inject Windows updates into one or more Windows ISO images (one edition per image at a time)" and
' "Inject drivers into one or more Windows ISO images (one edition per image at a time)" have been completely removed. The first of these sections
' was fairly useless since the first item in the menu basically has all the functionality of that section plus more. The latter section for injecting
' drivers has been merged into the code for injecting updates. Most of the steps for injecting updates and drivers are the same so it made sense to
' simply merge these into one routine. This has resulted in the removal of hundreds of lines of code. In addition, the two routines had very different
' status displays and in general just had a very inconsistant look. Now, there is a very consistant and unified look to the two operations.
'
' The routine called "Strip the time and date stamp from filenames" has also been removed because it is no longer needed with the above changes.
'
' In addition, the following minor improvements have been made:
'
' Routines that create bootable images previously would exclude any autounattend.xml file to prevent a situation where a user might accidentally boot
' from media that would start automatically installing Windows and wipe out their existing Windows installation. This has been changed so that the
' user is now informed of this situation and then asked if they want to exclude any autounattend.xml answer files.
'
' Updated a couple of messages for clarification. Also, for the recently added ability to scan DISM logs for errors, we have removed the question to
' the user to ask if that scan should be performed. After testing it has been determined that this should be mandatory because we want to avoid a
' situation where we end up with a useless final image without the user being aware of the situation.
'
' 8.0.0.34 - Feb 16, 2020
' Minor tweaks. Changed the working of text in a few spots and added color to a few areas.
'
' 8.0.0.35 - Feb 17, 2020
' Minor tweaks. There are a couple of options that need some explanation. The explanation for both is a little long. I changed
' these screens so that the use needs to type HELP to get a full explanation. This means that the screen will be much cleaner
' without all the extra information that the user probably only needs to see once.
'
' 8.1.0.36 - Feb 18, 2020
' Redid the routine to ask user if they want to configure a drive as GPT rather than MBR if the drive is larger than 2 TB.
' Prior to this update, the user was asked about this every time, now we only ask the user about this if the disk they selected
' is larger than 2 TB. Otherwise, we don't bother to ask.
'
' When creating a bootable disk, values had to be entered in MB and were displayed in MB only. Now, values can be entered in MB, GB,
' or TB, and are displaed in the most appropriate format as well.
'
' Compiled on QB64 ver 1.4
'
' 8.1.0.37 - Feb 19, 2020
' Changed the DIM statements for a few variables to DOUBLE to make sure we can hold the large numbers needed for storage space in the routine
' to create bootable media. Disk space is specified in MB and on very large disks these values can be large numbers.
'
' 8.2.0.38 - Feb 21, 2020
' In the routines for injecting updates and drivers as well as the routine to create a bootable ISO image that can include multiple editions and
' architectures, the ability to specify "ALL" for the index numbers has been added. This eliminates the need to have to display a list of the
' available editions to determine how many of them are present if you want to add all of them.
'
' Eliminated the DisplayIndices subroutine since it is no longer used in the program after recent changes.
'
' 8.2.0.39 - Feb 21, 2020
' Squashed a bug. In several places within the program, we would not detect that an ISO image was invalid if it contained an INSTALL.ESD
' rather than an INSTALL.WIM. This has been corrected.
'
' 8.2.0.40 - Feb 21, 2020
' Fixed a bug that was introduced in 8.2.0.39. When checking the validity of an ISO image, we were checking for a file called
' \x64\install.wim. This is incorrect. We should be checking for \x64\sources\install.wim. The result was that valid dual architecture
' images would be flagged as invalid. This issue has been corrected.
'
' 8.2.0.41 - Feb 23, 2020
' Up until now, the program was hard coded to run DISM and OSCDIMG from the default install locations. With the update, we query
' the registry for the install location and store that in a variable. As a result, we can run correctly even with the ADK installed
' to a non-default location. We can also warn the user if DISM.EXE or OSCDIMG.EXE are not found.
'
' 8.2.0.42 - Feb 23, 2020
' Further refinement of the ADK detection introduced in 8.2.0.41. Since some routines can be run without needing the ADK,
' rather than simply exiting the program, we warn the user that the ADK is not installed but then proceed to the main menu.
' If a menu selection is made that requires the ADK, we again warn the user that the ADK needs to be installed and return to
' the main menu rather thanrunning that routine.
'
' Also updated the branding. In some areas we referenced this program as simply "WIM" (Windows Image Manager) rather than
' "WIM Tools". This has now been made consistent throughout the program.
'
' 8.2.0.43 - Feb 24, 2020
' Performed some tidying up. Cleaned up the spacing before and after loop structures, added and reworded a few comments, declared every variable used in the
' program, changed variables that were declared as DOUBLE to LONG.
'
' 8.2.0.44 - Feb 26, 2020
' Modified the wording of a couple of messages to make them clearer.
'
' 8.2.0.45 - Feb 28, 2020
' Based upon an error encountered, we have add some error handling to the MountISO SUB procedure. If we cannot get a valid CDROM ID or drive letter for an
' image that we have mounted, we display an error message along with the CDROM ID and drive letter that we retrieved, and we then trminate the program.
'
' 8.2.0.46 - Feb 28, 2020
' Added some color to a message to highlight what we are asking for. This is in the routine for injecting Windows updates or drivers. If a dual architecture
' image is selected we ask if the user wants to add any of the x86 and any of the x64 editions. We are now highlighting "x86" and "x64" in the message to
' the user to make absolutely clear what we are asking for.
'
' 8.2.0.47 - Mar 01, 2020
' When the program is started, it will now add itself to the exclusion list for Windows Defender antivirus scanning. When you choose "Exit" from the main menu, it will
' remove itself from the exclusions. This will prevent a whole bunch of exclusions from being created if the filename or path from where the program is changed frequently.
' If the program is NOT terminated by selecting "Exit", simply run the program again FROM THE SAME LOCATION and choose "Exit" to remove the AV exclusion.
'
' 9.0.0.48 - Mar 13, 2020
' Major new version! I discovered that in addition to updating the main OS (INSTALL.WIM), the WinPE (BOOT.WIM) and WinRE (WINRE.WIM) components should also be
' updated. This program will now update those components when updates are injected. This required some major work to the program. In addition, this required a
' change to the layout of the updates in the folder supplied by the user. Please see the documentaion for details on how to organize your updates. Also revised
' the cleanup routine to properly handle the new folders and mounts for DISM used by the system.
'
' 9.0.0.49 - mAR 13, 2020
' The routine to inject drivers shares most of the same code with the routine to inject Windows updates. Before running the routine to inject updates we
' ask the user if they want to exclude an AUTOUNATTEND.XML answer file from the project if that file is present. At the very end of the routine where we build
' the final directory structure from which to create our ISO image, we check to see what choice the user made regarding excluding the answer file. However, before
' running the routine to inject drivers, we were asking about this. This in turn caused the logic for building the directory structure to massivly fail. This
' has been fixed by adding a single line of code to call the routine we neglected to include.
'
' 9.0.0.50 - Mar 13, 2020
' Minor fix - I neglected to updated the version number that is displayed with the 9.0.0.49 release. Correcting this.
'
' 9.0.0.51 - Mar 15, 2020
' Found and corrected a logic fault: We were grabbing the WinRE.WIM from the very first x64 and first x86 edition of Windows being processed and then
' using that file in all the other editions. The problem is that the WinRE.WIM is not present in a syspreped image. If a sysprep image was either the
' only or the first edition to be processed, then no WinRE.wim would be found. If a sysprep image was included but not as the first or only image, then
' a WinRE.wim from another edition would get inserted. We have fixed this logic.
'
' 9.0.0.52 - Mar 16, 2020
' Chased down some more bugs in the newly added routines. Much of problem I was chasing was the result, once again, of stale mounts open by DISM. Will need
' to remind users that interrupting program while a WIM is mounted by DISM can cause major difficulties and that cleanup will need to be undertaken. I am
' strongly considering greatly enhancing the cleanup routine so that it will do a true verification of any potential stale DISM mounts and report to the user
' if it cannot clean these up.
'
' 9.0.0.53 - Mar 16, 2020
' For the routines that inject Windows updates or drivers into a Windows image, we now check for, and attempt to resolve, any stale open DISM mounts. If the
' program cannot resolve the situation automatically, then we post a warning to the user suggesting a reboot and we exit the program.
' We also revised the cleanup routine (where user selects the cleanup option from the menu). First, it will try to clear all stale DISM mounts, not only
' those just in the directory specified. Second, in the event that a stale mount does exist, the revised routine shuld be far faster.
'
' 9.0.0.54 - Mar 17, 2020
' Reworked the DISM log file scan at the end of the routines to inject updates and drivers. Previously, it could take
' a couple minutes to process each log file. Now, we can process 25 log files in a matter of only a few seconds (3 seconds
' on my laptop).
'
' 9.0.0.55 - Mar 18, 2020
' Now that we can process a log file so quickly, rather than waiting until the routine ends, we are checking the log files at the end of each Windows edition
' that we process. If we do encounter an error, we inform the user, call the cleanup routine, and then take the user back to the main menu.
'
' 9.0.0.56 - Mar 18, 2020
' Found and fixed yet more logic bugs in the new WinRE and WinPE processing. Really hoping that this concludes this section.
'
' 9.0.0.57 - Mar 24, 2020
' Discovered that an error would getted logged in the DISM file if we tried to perform an update of a type that did not exist. As an example, suppose
' there are no servicing stack updates. An error would get logged when we tried to perform a servicing stack update. Added logic to intelligently skip
' any attempt at an update type that we don't have. NOTE: This was not a problem previously but with the new checking of the logs for errors it is a problem.
'
' 9.0.0.58 - Mar 24, 2020
' In the routines to inject updates and drivers, changed the wording in the status display to differentiage between an "image" and a Windows "edition".
'
' 9.0.0.59 - Apr 6, 2020
' In the routines to inject Windows updates, inject drivers, and Create a bootable Windows ISO image that can include multiple editions and
' architectures, we perform a check to see if the user is saving the project to removable media. If they are, we inform them that the
' project must be created on a HD. After that, we inform the user that they will be asked if the final ISO image should be moved after it is
' created. This would allow them to move the final image to removable media if they wish. The functionality of moving the final ISO image is
' not actually in the program. As a result, we have removed that part of the message so as not to mislead the user. We may add that capbility
' at a later time.
'
' 9.0.0.60 - Apr 8, 2020
' Neglected to update fileversion at start of program on build 59. This has been corrected. Added OSCDIMG logging for the routines that inject updates and drivers.
' Ran into an issue where OSCDIMG would log this error:
'
' ERROR: Could not open boot sector file "c:\project\ISO_Files\boot\etfsboot.com"
'
' Error 3: The system cannot find the path specified.
'
' Not sure if that error is because that file is read-only. For all routines that create bootable Windows ISO images, when we create the base image
' I have added a "/a-:rsh" to the robocopy commands clear the read-only, system, and hidden attributes from the files being copied to create the base image. In those routines
' where we are not creating a base image, I've added an "attrib" command to clear those attributes from the source files before we create the final image.
'
' In any case, encountering this error now is odd since I've never run into it before.
'
' 9.0.0.61 - Apr 12, 2020
' Modified program so that all temporary files will now be created in the TEMP directory. This will avoid visible files from appearing and disappearing
' in the location where the program was launched. To accomplish this, we also had to modify the behavior of the WimInfo.txt file that we gave the user
' an option to save. This was a text file that displays WIM information to the screen and could optionally be saved to that file. The previous behavior
' was that the file would simply be renamed from WimInfo.txt to WimInfo_Saved.txt right where it was located. Now that the file is created in the TEMP
' directory, we are simply moving it to the location where the program is located. It has also been renamed to WIM_Info.txt.
'
' 9.0.0.62 - Apr 12, 2020
' For the routines that inject updates and drivers, added the ability for the user to specify the final ISO image name being created. This will
' avoid confusion after a long running project and forgetting what the project was. Also found in flaw in the routine that checks for the existance
' of updates after the user specifies the update location. We were only checking for the existance of a Latest Cululative Update, but it's
' possible that user may only be applying another type of update(s). This has been corrected.
'
' 9.0.0.63 - Apr 13, 2020
' It appears that the registry key that I was searching for to verify that the ADK is installed on the system may change with the version of the ADK installed.
' A member of TenForums was kind enough to check for the existance of this registry key on his system since he is running a newer BETA of the ADK. Sure enough,
' that key does not exist on his system. I have determined a better registry key to check that makes a whole lot more sense and implemented it. This should be
' good even with newer versions.
'
' 9.0.0.64 - Apr 13, 2020
' In the routine that creates a VHD and deploys Windows to it, we are increasing the size of the EFI partition from 100MB to 260MB.  While this may not
' matter for a VHD it is nandatory for 4K advanced format drives. This places it inline with my procedures for all systems.
'
' 9.0.0.65 - Apr 14, 2020
' Increased limits on some arrays to 100 to allow up to 100 elements to be processed in a project. Also, for very long lines of code, these were split using
' an "_" to allow continuation on another line. This makes the code easier to read.
'
' 9.0.0.66 - Apr 14, 2020
' In reviewing some code I noticed that not all REDIM statements in the program had an AS TYPE clause. Corrected this to avoid potential problems.
'
' 9.0.0.67 - Apr 14, 2020
' A bug was introduced in build 9.0.0.63 when the change was made to the ADK path. This was a critical error that would cause the ADK to fail becaause the path
' was not properly built. Corrected and tested.
'
' 9.0.0.68 - Apr 15, 2020
' Corrected some text from "Do you want use at least one Windows edition".... to "Do you want to use at least one Windows edition"
'
' 9.0.0.69 - Apr 21, 2020
' When injecting updates, the status display treated the addition of of the Servicing Stack Update (SSU_ and the Latest Cumulative
' Update (LCU) as one item. To make the status just a little bit more granular, this has been changed so that the SSU update and the
' LCU update are broken up into two seperate lines.
'
' 9.1.0.70 - Apr 22, 2020
' Back in build 62 the ability to specify the final ISO image name in the routine that inject Windows updates as well as the routine
' to inject drives was introduced. This is now possible in the "Create a bootable Windows ISO image that can include multiple editions
' and architectures" routine as well. Since there have been a lot of changes recently, it's also time to bump the version number up
' from 9.0.0 to 9.1.0.
'
' 10.0.0.71 - May 30, 2020 (Updated version at start of this line, incorrectly listed this version as 9.2.0.71 in the actual release.)
' After much research and testing, found a way to create dual architecture images without the need for asking a user for a dual
' architecture image. All files needed to create a base image can be obtained from x64 and x86 media; some files from the x64 media,
' other files from the x86 media. The one exception is that two different files, both named "bcd" are needed from the dual
' architecture media. The breakthrough came in that a way has been found to recreate those files from plain text which allows the
' creation of those files on the fly from within this program.
'
' 10.0.1.72 - June 10, 2020
' The first time updating the final release of Windows 10 2004 I found that errors were being logged by DISM. After the project
' was done, Windows installed just fine, so these may not be fatal errors. I am investigating this. In the meantime, I have
' updated how errors are handled. Rather than stopping ast the first error, the status screen will show a message indicating
' that an error was detected. We save all errors to a seperate log file that stores nothing but errors as well as what log
' those errors came from and a timestamp. When the routine is done, we remind the user that errors were detected and point them
' to this log file.
'
' 10.0.2.73 - June 10, 2020
' Reword a couple of places in the program where we display text that says "working on image"... Changed this phrase to "working on edition" for accuracy.
' In addition, where a user has an option to update multiple editions in an image, if a user responded "ALL" when asked what editions to update, there
' was a bit of pause that might make one wonder if the response supplied was received. Changed this to immediately display a status message as soon as a
' response is provided by the user.
'
' 10.0.3.74 - July 14, 2020
' When injecting updates into Windows editions, after each edition is updated we parse the DISM log file to see if any errors were reported. If so,
' we display a message that errors were detected and inform the user that more information will be displayed after the routine is done running, in
' other words, when all editions have been updated. Prior to version 10.0.1.72 as soon as we detected an error we would abort the program. In
' version 10.0.1.72 we changed that behavior because a bug that I confirmed with Microsoft is causing errors but these can be ignored. So, the only
' change made in this version is to inform user that they can check the log file as soon as the program displays the message indicating errors.
' There is no need to wait until the program is done.
'
' 10.0.4.75 - July 17, 2020
' For the routine that creates a VHD and deploys Windows to it, we asked for the index number to use before we gathered other information. While this
' worked, the ordering of the information gathering just seemed out of place. Also improved the messaging to the user in this section of code.
'
' 10.0.4.76 - August 3, 2020
' Very minor update. For some reason the words "partition" and "partitions"  were mispelled in several locations. This has been corrected.
'
' 11.0.0.77 - August 20, 2020
' Major new revision: Added a whole new capability; the program can now add boot-critical drivers to the WinPE and WinRE images.
'
' 11.0.1.78 - August 21, 2020
' Minor change. For the routine the displays WIM information, if the user chooses to save a copy of the output to a file we now
' display a message to let the user know the full path and filename of the saved file. This makes it clear for the user exactly
' where to find the file. In addition, if the user elects to not save a copy of the output to a file, then rather than displaying
' a message that says "done", we will simply return to the main menu.
'
' 11.0.2.79 - August 28, 2020
' There are several places in the code where we create an ei.cfg which is placed in the \sources folder of the Windows media.
' This file is no longer needed. Rather than remove that code entirely, those sections have been commented out for now just
' in case any reason should ever be found to reinstate that code.
'
' 11.0.3.80 - August 31, 2020
' A change in direction from the last update. Rather than completely disable the creation of the ei.cfg file, we are now asking
' the user if they want to create an ei.cfg file.
'
' 11.0.3.81 - September 4, 2020
' Minor change. The main menu item number 10 had the line indented one space too far. corrected this.
'
' 11.1.0.82 - September 9, 2020
' This update adds the ability to inject generic files into the WinPE image (boot.wim). This can be helpful for purposes such as
' the addition of a script to create disk partitions (useful for creating the WinRE partition last). In addition, the recently
' added code to allow boot-critical drivers to be injected introduced a bug where the x86 version of that code would run when
' the user was adding Windows updates, not boot-critical drivers. This has been fixed.
'
' 11.1.0.83 - September 10, 2020
' Very minor update that does not affect functionality at all. The prompt for making a selection from the main menu was crowded
' right up against the menu above it. Moved the prompt down one line. It just looks a little better.
'
' 11.1.1.84 - September 11, 2020
' Fixed a bug. The recently added code to add generic files to the boot.wim file used a generic variable named "x" to keep count of
' the number of files to be copied. Unfortunately, this was done within another loop that was already using that same variable name.
' As a result, that variable was being incorectly altered. To fix this, we changed over to a different unused variable (z).
'
' 11.1.2.85 - September 11, 2020
' For the routine that reorders entries in the ISO image, I found that if the source and destination are the same directory, the original
' file is not replaced. There is no error, the symptom is simply that the reordering of Windows editions never happens. Investigation
' shows that this was due to the original file still being mounted, making it impossible to overwrite the file. To correct this, the
' dismounting of the original file was moved up to just before we create the new file. In addition, changed a couple of messages to the
' user in this section. The program displayed messages with the word "image" where "index" would be more accurate. Corrected this.
'
' 11.1.3.86 - October 31, 2020
' With the release of Windows 10 20H2 there appears to be a bug with applying the Latest Cumulative Update (LCU) to the Windows PE
' image (boot.wim). When the LCU is applied to boot.wim and Windows is installed using an image that incorporates that update,
' Windows fails to detect the HDs during setup and displaying a message that a media driver is missing. To work around this issue
' for now, we ask the user if they want to skip SSU and LCU updates for Windows PE. The user can elect to skip these updates which
' will prevent the problem from happening. Note that any user files to be added or deleted from the WinPE image are unaffected.
' Even if user skips application of SSU and LCU updates, these files will still be properly processed.
'
' In addition, if the routine to inject updates finds errors in the log files, and that routine is run a second time without first
' exiting the program, the program will indicate that errors were found even if none exist. This is because a flag that indicates
' that errors were found does not get reset when the routine is first started. This has been corrected.
'
' Changed references to "Disk ID" in the program to "disk number" as this terminology could be confusing, especially in the context
' of how the disk IDs were displayed where they are simply presented with disk numbers.
'
' 11.1.4.87 - November 2, 2020
' The routine to change the NAME and DESCRIPTION metadata was not working. The problem was tracked down to a wrong path being stored
' in a variable that is supposed to contain the full path to the ImageX.exe executable. Rather than specifying Imagex.exe it was
' referencing DISM.exe. This has been corrected.
'
' 11.1.5.88 - November 24, 2020
' Microsoft has released a new type of update for the Latest Cumualtive Update (LCU) known as a "GDR-DU" (General Deployment
' Release - Dynamic Update). This file is distributed as a >CAB file rather than as a .MSU file as the LCU has always been
' released previously. Since this program checks the folder where the LCU is located for a .MSU file, the program has now
' been modified to look for either a .MSU or a .CAB file. Either will be considered acceptable. The commands that actually
' inject the updates into the Windows image dis not need to be modified since they handle .CAB files as well as .MSU files
' without any changes at all.
'
' 11.1.5.89 - November 26, 2020
' No changes to the code other than updating the build number, the release date and the release notes. The purpose of this
' build is simply to recompile the program using the Nov 4, 2020 development build of QB64.
'
' 11.2.0.90 - December 9, 2020
' Several improvements have been made to the routine that creates bootable media from an ISO image or to update already
' existing boot media. First, it is no longer necessary for the user to supply drive letters for all the partitions
' being created. The user now has the option of letting the program automatically assign drive letters, however, if they
' would still prefer to assign drive letters manually, they can do so. In addition, the volume names that were assigned
' to the partitions all started with "USB". This has been changed because this routine can be used with non-USB based
' media such as SD cards and HDs attached via other means than USB.
'
' A few other minior tweaks were made as well. The wording of a few messages were slightly changed. When user enters the
' desired size for partitions, the display has some coloring added to it to better highlight the values provided and for
' which partitions responses are still pending.
'
' Finally, starting with this build, we move to compiling using QB64 December 7, 2020 development build.
'
' 12.0.0.91 - December 12, 2020
' Elevating version number to a new major version number in recognition of all the major enhancements made recently.
'
' Yikes! Neglected to remove some code for testing. This has been corrected. Also corrected a spelling mistake that was noticed in a comment.
' When creating a boot disk, if the user chooses to perform a refresh operation, we now maintain the volume label. This is important
' if you rely on the volume label for anything. As an example, I use a script that references the volume label rather than relying on
' drive letters which can change. Also, when performing a wipe operation, rather than a prefined set of volume labels the user is now
' prompted for volume labels.
'
' When creating a bootable Windows disk with additional partitions, if the user elected to enable BitLocker and BitLocker failed,
' for example, if the same password was not entered twice, the program would proceed with no warning to the user and no chance to
' try again. This has been corrected. The user now has 3 trys. If it still fails, the program will continue gracefully but will
' warn the user that BitLocker will need to be enabled manually.
'
' Finally, the dimensioning of one variable that was used only in testing but is no longer needed was removed
'
' 12.0.1.92 - December 12, 2020
' The version info had a typo - it was listed as 10.0.0.91 but should have been 12.0.0.91. This has been corrected. Also, there
' was a section of code that saved data to the clipboard. This has been changed so that the program will no longer alter the
' contents of the clipboard.
'
' 12.0.2.93 - December 14, 2020
' A solution to the previously identified problem of updates to the WinPE image (boot.wim) causing setup to fail has been identified
' and implemented in the code. Microsoft's own documentation was to blame. The documentation neglected to note that for the 2nd index
' in the boot.wim file, when exporting with DISM, a "/bootable" switch should be used. Oddly, prior to the October 2020 updates, the
' lack of this switch did not cause any difficulties. However, from October onward, this caused problems. Thi switch has now been
' added. In addition, a simple bypass for the code asking a user if they wanted to skip updates for WinPE has been implemented. This
' is much safer than trying to address all the sections of code where this is implemented, and allows us to easily restore that
' functionality if needed by simply commenting out 2 lines of code. If all is well, we may consider cleaning up this code at a
' future time.
'
' 12.0.2.94 - December 15, 2020
' No code changes at all in this update (other than the version number and this note). This build is simply a reversion back to QB64 1.4
' for compiling. It's not that any problems were found in the development build, but just want to ensure that I am on a stable release
' and not risking any problems by being on a development build.
'
' 12.0.3.95 - December 18, 2020
' I was bored so I made a few minor tweaks. There are no functionality changes with this update. I simply reworded a few prompts
' and added a few colored highlights for emphesis here and there. Noticed that I used the word "images" where I meant "editions"
' so I corrected that.
'
' 12.0.4.96 - December 18, 2020
' Again, no functional changes. Just cleaning up and rewording some messages to something I like better and adding a few more color highlights.
'
' 12.1.0.97 - December 19, 2020
' There is Microsoft issue that has been around for a while (quite a few months) that cause DISM to log false errors. Because of this issue,
' messages are being displayed to the user which would cause undue concern. I've disabled the code that shows these status messages for now.
' This is a bit frustrating because, if there are any real errors to be aware of, the user will not know until the update routine is done
' and that could be a considerable amount of time. Related to this, the logging has now been revamped. We still have the log file called
' ERROR_SUMMARY.log that contains all errors, but we also create a new log file called SANITIZED_ERROR_SUMMARY.log that sanitized the
' log and removes all of the known false errors. Comments are present in the code to easily undo the bypassing of the error staus messages.
'
' 12.1.1.98 - December 20, 2020
' Squashed a rather nasty bug discovered while testing another function. In the option to create an ISO image from one or more editions of
' Windows, it was discovered that if a dual architecture image was being created, the entire x86 folder would be missing from the final image.
' Analysis: There one section of code that runs if dual architecture image is available. This section of code makes use of an existing dual
' architecture image to copy all the files that we need to create a base image for a dual architecture project. However, if all the images
' in the project are single architecture, but we have a combination of both x64 and x86 editions of Windows, then we need to take a different
' series of steps to create the base image. The problem was that after the first section of code was run, we should have had a simple GOTO
' statement that would skip over the second section code. The idea is that we only need one or the other section of code depending upon the
' circumstances. However, this GOTO statement was missing. As a result, we ended up running both secontions of code with bad consequences.
' Naturally, this has now been corrected.
'
' 12.1.2.99 - December 20, 2020
' Performed a little bit of revamping for the feature to allow reordering of Windows editions wihin an image. Added help to the prompt
' asking for the index order because it may not have been readily obvious that indicies, including a range of numbers can be entered in
' either ascending or descending order or a combination of both. Cleaned up the text for a few more messages.
'
' 12.2.0.100 - December 21, 2020
' Performed a few changes to the routine that reorgaized the Windows editions within an ISO image. We always have allowed the user to
' specify the same folder as both source and destination. This causes the original file to be overwritten with the new file. The problem
' with this logic is that if the user does not specify all indices, then some editions will be omitted and we would loose Windows editions.
' Just to add a little bit of safety, if the user specifies the same source and destination folder, we will now inform them that the
' original file will be replaced with the new one.
'
' In addition, this routine creates several temporary folders and a file in the destination folder. To be on the safe side, we now check
' to see if any of these already exist. If any are already present we inform the user and ask them if it is okay to erase these.
'
' 12.2.1.101 - January 1, 2021
' No new features. Cleaned up / reworded some messages and performed some minor touchups. For the moment we have no major updates
' to address or bugs to squash so we are taking the time to make some very minor adjustments. Found a number of references to a
' variable named ArcTag$. Nowhere in the code is this variable ever set to anything so I have removed all references to it. Just
' in case this proves to cause any difficulties, go back to version 12.2.0.100 to review where that variable was used. Added some
' more comments to just a few variable definitions to explain there usage. This is a very low priority so I may continue to
' document the usage of variables here and there over time when I am bored. The release notes for version 12.0.0.91 contained
' the words "testcode" (with a space between words). I modified that wording because I use that exact phrase when placing code
' for testing in the program. This way, when I search for that phrase in the future, I won't find it when no code for testing is present.
'
' 12.2.2.102 - January 1, 2021
' Added a few more comments to variable definitions
'
' 12.2.3.103 - January 7, 2021
' A little more cleanup of messages, added so colors for emphesis. Today we focused on the routine to create a Virtual Disk.
'
' 12.2.4.104 - January 13, 2021
' There was a problem related to updating Windows PE (boot.wim). We implemented a workaround in a sub called Skip_PE_Updates_Check.
' This section asks the user if they want to skip WinPE updates and then skips those updates if the user wants to do so. Note that
' if the user has files that should be added or removed to the boot.wim (not Windows updates), those will still be applied. We had
' believed that we had found a solution to this issue (adding a /Bootable to the export command for index #2 of the boot.wim) but
' it turns out that this does not resolve the issue in all circumstances. Since the problem is now reoccurring again we are
' reactivating that code that we had disabled starting in build 12.0.2.93 on Dec 14, 2020.
' Also related to that same section of code, it was discovered that a "/Bootable" switch when exporting WinPE (boot.wim) Index #2 in
' the routine that injects updates into Windows editions was missing. That has been resolved.
' To summarize, the workaround code has now been re-implemented.
'
' 12.2.5.105 - January 14, 2021
' Added 2 more messages to the list of messages stripped off from the Sanitized Error Summary log.
'
' 14.0.0.106 - Jan 15, 2021
' FINALLY - I believe that we have a definitive answer to the Windows PE (boot.wim) upgrade issues. It's taking a fair bit of code to
' rectify the situation and we'll be testing the fix for at least a day or two, but I'm pretty sure we have it this time around. Since
' we now have a fix, we are disabling the Skip_PE_Updates_Check routine which was designed to ask a user if they wanted to skip the
' WinPE updates as a workaround for this issue. The code is still there and can be re-enabled by simply commenting out 2 lines.
' Please go to the Skip_PE_Updates_Check routine and see comments there for details.
'
' 14.0.1.107 - Jan 18, 2021
' Testing indicates that my thinking I had the definitive solution to the issue I described in the past several updates was spot on.
' Unfortunately, I had one small glitch. In one line of my code I built a command line to be executed but never executed that command.
' That problem was in the vicinity of the code I needed to fix this issue and it was pure chance that testing that fix revealed this
' flaw. That issue has now been resolved.
'
' 14.0.2.108 - Jan 19, 2021
' Added 2 more messages to the list of messages stripped off from the Sanitized Error Summary log.
'
' 14.1.0.109 - Jan 22, 2021
' Changed the order of menu items and reordered those major code subsections so that they appear in the code in the same order as
' reorganized menu items. The code sections really didn't need to be moved, It was just extremely easy to do and I thought it
' would be nice to keep it organized that way. As for the menu; it had never really occurred to me that the tools not really
' directly related to WIM management appear in the menu between WIM management functions. These tools have now been moved so
' so that all WIM management functions are grouped together and the other helpful, but unrelated tools, appear after these.
'
' 14.1.1.110 - Jan 26, 2021
' Variale type values assigned from FREEFILE command should be of type LONG but I was using the INTEGER type. There
' was no error observed but I thought best to correct this while I noticed it. This was noticed while addressing this
' issue: In the process of running this program on a system that had a BitLocker encrypted drive, AND that drive was
' also currently in a locked state, I discovered that the routine to find the next available drive letter would think
' that the letter assigned to a BitLocker protected drive  was available. This was corrected by performing a further
' check on any drive letter we initially believe to be free. If the _DIREXITS function indicates that a drive letter
' does NOT exist, we run a "manage-bde -status D:" (or whatever drive letter) command and parse the output. If that
' command includes "could not be opened by BitLocker" in the output, then we know that the drive really does not exist.
'
' 14.1.2.111 - Feb 3, 2021
' For the x64 code only, when exporting the boot.wim index #2 in the routine that injects Windows updates, we have
' neglected to hide the output of the DISM command performing the export. This causing some ugly output on the status
' screen. If may be very brief depending upon the speed of the computer, but it looks rather ugly and even on a fast
' machine causes a brief series of flashes and poor looking output. This has been corrected.
'
' 14.1.3.112 - Feb 4, 2021
' In a section of code where we are assigning drive letters to partitions, we open a file after determining the first available
' free file number and storing the value in the variable "ff". In that section of code, rather than referencing "LOF(ff)" we were
' referencing "LOF(1)". That happens to work because ff will likely always be equal to 1, but with future changes to the code, we
' cannot guarantee that this will be the case. This has been corrected to "LOF(ff)".
'
' In the routine that created a VHD, deploys Windows to it, and then adds it to the boot menu, we referenced a variable named
' IndexVal before it was set to anything. This was an error. At that point in the code, rather than reference a variable, we
' should simply be using a fixed value of "1". This has been corrected.
'
' 14.1.4.113 - Feb 4, 2021
' Revamped the entire routine to create a VHD, deploy Windows to it, and add it to the boot menu.
'
' 14.2.0.114 - Feb 25, 2021
' If a Windows edition has a pending operation, this will prevent DISM Image Cleanup operations from being performed. The program
' has been updated to inform the user if at least one Windows edition has been found to have pending operations. Note that this
' typically happens if NetFX3 is enabled on the Windows edition. If this happens, the user should update Windows editions that do
' not yet have NetFX3 enabled on them. After the other updates are applied or the install.wim has had the cleanup operation
' performed on it, NetFX3 can then be enabled.
'
' In addition to the above some very minor updates were made. For example, prompts for certain information to be supplied by the
' user would display a question mark where it was not needed at the end of the prompt. This has been corrected.
'
' 14.2.0.115 - Feb 25, 2021
' There are no code changes in this version other than to update the version number and to remove "_DEST _CONSOLE" since this is no longer
' needed in QB64 1.5. The main purpose of this version is simply to recompile the program using the new version 1.5 release of QB64.
'
' 15.0.0.116 - Mar 5, 2021
' Major new release. The initial work that was started in build 114 to check for pending operations on a Windows edition has been
' greatly enhanced. While we still display a message to the screen that we have detected a pending operation, we have also made
' major changes to the code that allow the original Windows images and editions to be tracked through the update process so that
' when a pending operation is detected we can log exactly what the original file and index that has the pending operation. This
' information is then saved to a log file so that the user can review the details.
'
' A few other more minor tweaks have also been implemented. For example, the operation to cleanup project folders has been
' enhanced to display the results in a clearer manner more consistent with the rest of this program.
'
' 15.0.1.117 - Mar 7, 2021
' In the section of code for injecting updates, even though I disabled the display on screen of error warnings there was a section of the
' message still being displayed refferring users to an error log. I have commented that out for now and may eventually remove it entirely.
'
' 15.0.2.118
' When adding updates to a Windows image, it's possible that no errors may be generated at all, even the false positives that we filter out
' in the program, depending upon what options the user selects. In this case, no file named SANITIZED_ERROR_SUMMARY.log will be created. In
' order to make the program a little friendlier, we check for the presence of this file at the end of the routine. If no such file exists,
' we no longer suggest to the user to take a look at that file to make sure all is good.
'
' In addition, in the process of looking at the above issue, it was noticed that there was a spot in the code where we reference "eof(1)",
' however the file was not being opened as #1, but was being referenced by the filenaumber assigned by the "freefile" function. The code
' was working fine because "freefile" happened to be #1 every time, but just in case any future code changes alter this, we have corrected
' this so the "eof" now references the variable that "freefile" is associated with.
'
' 16.0.0.119
' Added the ability to add an autounattend.xml file to the image in addition to other updates. The best way to accomplish this was by
' completely re-architecting the directory structure for the Windows updates. Since an answer file does not get placed in either the
' x64 or x86 folders we needed to make some changes. As a result, the program will no longer ask for the location of both x64 and x86
' updates. It will only ask for the location of the updates in general and will require that the updates folder have a seperate x64
' and x86 subfolder for the updates. Note that if only one architecture type is being processed, there is no need for a folder for
' the other architecture type. Within the updates location, in addition to the x64 and x86 folders, a "Answer_File" folder should be
' created. Within that folder will be the "autounattend.xml" if it is intended to be added to the project.
'
' Because this update fundamentally changes the structure of the updates folder that the user needs to supply, we are assigning a major
' new version number and including in the release notes information regarding this change.
'
' 16.0.1.120 - Mar 17, 2021
' The routine used to determine if the program was being run elevated has been reworked so that it is much shorter and now does not
' require the use of any variables whatsoever. In addition, it was moved to the very beginning of the program before we even dimension
' any variables.
'
' In addition, at the very start of the program, we have placed a CLEAR statement to wipe out all variables and arrays. After any routine
' is completed, rather than return to the main menu, we return to this point now. This keeps things tidy and limits memory usage. Granted,
' this program does not use much in the way of resources, but it seemed like a good thing to do.
'
' In the course of adding this CLEAR statement, a bug involving interaction of the CLEAR statement with a console was discovered and
' reported to QB64.org. A fix was developed that very same day (highly impressive!). This fix is implemented in the March 17, 2021
' development build of QB64. As a result, it is important to use that build or newer when compiling this program in the future.
'
' 16.0.2.121 - Mar 19, 2021
' Sloppy work - In the last build, a single line of cose was accidentally deleted which caused the routine to inject updates to fail
' to run. When selecting that option from the menu, an error would be returned saying that an invalid menu selection was made.
' This issue has been corrected.
'
' 16.0.3.122 - Mar 22, 2021
' No changes in functionality. Simply reworded a few messages to make their mening clearer.
'
' 16.0.4.123 - Mar 30, 2021
' In creating a boot disk from a multi-edition image previously created using this program, it was found that we were out of disk space
' on the intial FAT32 partition. As a result, we are updating the program to create a larger FAT32 partition. We will now create that
' partition with a size of 2.5GB rather than 2GB. While making this change it was also noted that one instance of the word "partition"
' was misspelled as "partion". This has been corrected.
'
' 16.0.5.124 - Apr 10, 2021
' Bug fix: When injecting updates into Windows editions, we output the results of command to a temporary file. That file is then
' parsed for the occurence of the stings "x64 Editions" and "x86 Editions". Also contained in that file is the name of the file
' itself. If that file name happens to contain "x64 Editions" or "x86 Editions" in exactly that case, it will trigger a false
' positive in our tests causing the wrong actions to be taken. This has been resolved.
'
' 16.0.6.125 - Apr 12, 2021
' Extremely minor change. When the program is run and a selected operation is completed, for example, the process of applying
' updates to Windows editions, the program will prompt you to press any key to return to the main menu. When a key is pressed
' there is no immediate visual indication that the keypress was accepted. On a fast computer there is only a momentary delay,
' but I believe that on a slower system there may be several seconds delay as everything is initialized. Since the program is
' purposely designed to buffer user input, we don't want the user to press further keys thinking that nothing has happened as
' these will be stored and played back. To remedy this we no immediately clear the screen and display a message to indicate
' that initialization is in progress.
'
' 16.1.0.126 - Apr 14, 2021
' While working with certain thumb drives that are supposed to be very fast, it was noted that the performance had slowed down
' to a crawl. By writing all zeros to these drives using the "CLEAN ALL" command in DISKPART, performance was restored. A
' thumb drive does not have TRIM support, however, in theory, there should be some sort of garbage collection capability
' which should do something similar.
'
' Since exFAT is the preferred format for larger thumb drives, it's possible that the garbage collection methods used by
' these thumb drives may simply not handle NTFS. As a result, for the routine in this program that creates bootable media, we
' have added an option to allow formatting of all partitions other than the first partition as exFAT rather than NTFS. The
' user can choose which they prefer. Note that at the start of the routine a variable named "UserCanPickFS$" is set to "TRUE"
' by default. By changing this in the code to "FALSE", this new functionality can be disabled and the program will then always
' use NTFS. This is considered experimental at this time since it's not know how well using exFAT on the boot media will work
' on all systems during Windows installation. This will be tested on a number of systems in the coming days.
'
' Starting with this release of the program we are moving to QB64 1.51 April 9, 2021 Development build for compiling.
'
' 16.1.1.127 - Apr 14, 2021
' Following some routine testing after the changes made in 16.1.0.126, it was discovered that a boot drive created by the program would
' not boot on a BIOS based system. There is a section of code where a determination is made as to whether the drive should be configured
' as MBR or GPT. Normally, the drive should be configured as MBR but due to a logic bug we never execute the command to convert the drive
' to MBR. This has been resolved.
'
' 16.1.2.128 - Apr 15, 2021
' One more minor fix. While formatting a slow SD card on a slow computer, BitLocker failed to initialize in the 30 second timeout that
' alotted. Increasing the timeout to 90 seconds. This should not slow the program down normally because we will move on once the
' BitLocker initialization is completed.
' Also cleaned up a few messages to make them easier to understand.
'
' 16.2.0.129 - Apr 18, 2021
' For the routine that exports drivers from a system, we've added one new feature: After exporting drivers a batch file will
' be created in the export directory for the purpose of installing all the exported drivers. The needs to simply run the
' batch file to install the drivers. Since this batch file is self-elevating, there is no need to remember to run elevated
' or to have to remind the user to do so.
'
' 16.2.1.130 - Apr 18, 2021
' Very minor revision to the batch file created as described in the release notes for 16.2.0.129. We have added text at the
' end of the routine indicating that a reboot may be needed. Also, we are hiding all the messages that would normally
' get displayed when drivers are exported as well as when they get installed by the batch file. It just looks so much
' cleaner without all those messages.
'
' 16.2.2.131 - Apr 19, 2021
' Found a rather bug that affected the routine that injects Windows updates, Windows drivers, and Boot Critical Drivers. For
' each of these routines, if we were updating x64 editions of Windows only and the updates or drivers folder did not contain
' a \x86 folder, the routine would display an error stating that no x86 updates or drivers could be found even though we
' were not processing any x86 editions. This issue is now resolved.
'
' 16.3.0.132 - Apr 19, 2021
' For the routine that creates a bootable disk, the manner in which the user was made to select the disk ID for
' the disk to be made bootable was very clunky. This entire portion of the routine has been completely reworked.
' It was now much easier to understand and navigate. In addition, error checking has been added so that a user
' cannot select a disk ID that does not exist. Finally, the program previously would not accept a disk ID of 0.
' However, the OS disk is not necessarily always disk 0, so disk 0 should be considered a legitimate response.
'
' 16.3.1.133 - Apr 19, 2021
' Changed a couple of messages for clarity.
'
' 16.3.2.134 - Apr 28, 2021
' Changed one message to make it look better. Text went only half way across the screen and then continued on the next line.
' It just look a bit sloppy. This is also the first version to be compiled on QB64 1.51 Apr 25, 2021 development build.
'
' 16.3.3.135 - May 22, 2021
' No program changes. The only change is that the program is now being compiled using QB64 v.1.51 May 22, 2021 Dev Build.
'
' 16.3.4.136 - June 5, 2021
' Updated the code that recreates BCD files on the fly to generate the new files from the 21H1 media.
'
' 16.3.5.137 - June 8, 2021
' This updates adds some enhancements to the AV exclusions. It was found that merely excluding the program
' executable was not sufficient. This update adds exclusion for the destination folder in the routines that
' inject updates, drivers, and boot-critical drivers. The folder name is logged to a temporary file. That
' way, if the program is terminated before the exclusion can be removed, when the program is started again
' it will detect that the exclusion was not removed and will then remove it when the program is started.
'
' 16.3.6.138 - June 15, 2021
' Changed the order of release notes so thatthe latest release is at the bottom rather than the top. With the
' release notes at the top, it was becoming a pain to find the latest note to add a new release note each time.
' By reversing the order of the release notes, it will now be much easier to update the notes.
'
' 16.4.0.139 - June 17, 2021
' For the routine that creates a bootable drive, we have added error detection for the robocopy commands that
' copy the data to the drives. This allows us to warn the user if an error is detected.
'
' In the process of updating this code, it was also found that the handling of the autounattend.xml file, if it exists
' on the source, was not being handled properly. This has been corrected.
'
' In the routine to create bootable media, we have a variable called "UserCanPickFS$". If set to "FALSE" then all partitions other
' than the first will automatically be created as NTFS. Set to "TRUE" it will allow the user the select whether to use NTFS or
' exFAT. We had this set to TRUE but we have now changed it to FALSE as this should be something that is almost never needed.
'
' Finally, starting with this build, we are compiling on QB64 1.51 June 14th 2021 Development Build.
'
' 16.4.1.140 - June 17, 2021
' Performed some additional cleanup of the routine to create bootable media. The displayed output is now much cleaner.
'
' 16.4.2.141 - June 17, 2021
' No changes to functioality, simply cleaned up some messages in the sections update in the last few builds.
'
' 16.4.3.142 - June 17, 2021
' Just some more housekeeping claifying a prompt to user to make it easier to understand.
'
' 16.5.0.143 - June 22, 2021
' For long running routines such as injecting Windows updates or drivers into an image, we have added the ability to
' specify whether the program should return to the main menu after running, or perform a system shutdown. To perform
' a system shutdown, the user should create a file named "Auto_Shutdown" (case does not matter) on the desktop. The
' status screen will update the status of auto shutdown after each step prior to a new step in the updare process
' being started. The user can create this file or delete / rename it even while the program is running.
'
' 16.5.1.144 - June 25, 2021
' Just a little minor change today. For the routine that creates or refreshes bootable media, we have provided a more
' friendly closing summary when the routine is done to reflect the options chosen by the user.
' Also fixed one other very minor issue. There are two messages that get displayed, one right below the other. The
' wording is exactly the same except for the number at the end of the message. One of these messages ended with a
' period, while the other did not. This just looked a little sloppy. It has been corrected by removing the period.
'
' 16.5.2.145 - July 6, 2021
' No change in functionality. Recompiled on QB64 1.51 July 6, 2021 Dev Build.
' Added a comment line after setting the variables that hold the current version and release date that notes the version
' of QB64 that this build is compiled on. This will simply make it a little easier to locate the QB64 compiler version
' used in the future.
'
' 16.6.0.146 - July 9, 2021
' We have revised the manner in which autounattend.xml answer files are handled. For all routines we will now exclude the
' autounattend.xml answer file if it exists. There are 2 exceptions to this rule:
'
' 1) For the routine to make or update a bootable drive from a Windows ISO image, we will ask the user if they want to
' exclude an answer file if one exists.
'
' 2) For the routine that creates a bootable ISO image from files in a folder we will INCLUDE the answer file if it is
' present since the user can simply delete it from the folder if they do not want it to be copied.
'
' Note that for the routine that injects Windows updates, it will exclude answer files from the original sources, but if
' you include an answer file in the "Answer_File" subfolder of the folder that you specify as the Windows updates location,
' this answer file will be copied.
'
' Finally, we have also renamed a few items on the main menu. We previously referenced ISO images where it would be more
' accurate to reference Windows editions.
'
' 16.6.1.147 - July 15, 2021
' Very minor update. Simply updated the comments in the first few lines of the program. Modified one reference to Windows 10
' to include Windows 11. In addition, removed a comment that referenced an older build of QB64.
'
' 16.6.2.148 - July 20, 2021
' Found that the routine to create the ei.cfg file if the user elects to create one was misplaced. This could cause that file
' to not be created even if the user elected to create it. This has been corrected.
' Corrected a typo in a path name. We had a path that should reference "ISO_Files" but it was spelled "xISO_Files".
' Corrected another spelling error where the word "least" was spelled "lease".
'
' 16.6.3.149 - August 2, 2021
' Added the ability to pause program execution. For the routines that inject updates, drivers, or boot-critical drivers,
' execution can use quite a few resources, especially disk resources. When viewing the status on the screen, as progress
' goes to the next item in the list, we will check the desktop for a file named "WIM_PAUSE.txt". If such a file exists,
' execution will be paused until that file is deleted or renamed. If execution is paused a flashing message will be
' displayed to make the user aware of this.
'
' 16.6.4.150 - August 12, 2021
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
' 16.7.0.151 - August 16, 2021
' For the routine that will create bootable Windows installation media, we have made some improvements to make the user
' experience better. First, we create an empty text file with a unique name to the first two partitions when a "Wipe"
' operation is performed. In the future, if the user wants to perform a "Refresh" operation, we will search for these files.
' The purpose of this is to detect if more than one bootable drive created with this program exists. If more than one such
' drive is connected, then we need to ask the user for the drive letters to be refreshed. Otherwise, if only a single such
' drive is connected, we automatically detect the drive letters to be updated so that we don't need to bother the user for
' those details.
'
' 16.7.1.152 - August 16, 2021
' Added a line to the main menu to indicate the edition (Dual Architure or x64 Only). The name of the source file and the
' resulting executable have also been changed to clearly differentiate between the editions.
'
' Undid a recent change to the self-elevating code. It turns out that this was causing the code to actually NOT self-elevate.
'
' 16.7.2.153 - August 17, 2021
' After further investigation into the self-elevating code issue, we now have a more robust solution. The comments in that
' section of code explain how it works.
'
' 16.7.3.154 - August 17, 2021
' For the routine that creates bootable media, we have increased the amount of time that we wait for BitLocker to initialize
' a disk if the user chooses to create BitLocker encrypted partitions. This should ensure that even very slow media, such as
' some thumb drives, have enough time to initialize. In addition, added a block of text to better guide the user when they
' choose to create additional partitions.
'
' 17.0.0.155 - August 26, 2021
' This is a major new release. Comprehensive help has been added to the program.
'
' 17.0.1.156 - August 26, 2021
' Added comments to the new help system in order to make the code easier to read and locate specific
' help topics within the code.
'
' 17.0.2.157 - August 31, 2021
' Performed some tidying-up of the text in the new help section.
'
' 18.0.0.158 - September 2, 2021
' Another new major version. For the routines that inject Windows updates, drivers, or boot-critical drivers, the program can
' automatically generate a script file complete with comments. The comments make it easy to manually modify or "tweak" the
' script file. These script files can then be played back automatically by the program.
'
' 18.1.0.159 - September 2, 2021
' Added the ability still create a script without actually having to carry out an injection of updates, drivers, or boot-critical
' updates. Also improved the look of the screen asking if user wants to perform any scripting operations.
'
' 18.1.1.160 - September 2, 2021
' Added a hint to the status screens to remind user that they can use AUTO_SHUTDOWN.TXT and WIM_PAUSE.TXT files on desktop
' to perform an automatic shutdown or pause the program execution.
'
' 18.1.2.161 - September 3, 2021
' Bug fix - the final ISO image created by the routines to inject Windows updates, drivers, and boot-critical drivers was
' lacking the ".ISO" extension on the filename. This has been resolved. Also fixed a bug where script file may fail to be
' moved to the program directory after it is created.
'
' 18.1.3.162 - September 3, 2021
' Bug fix - In the help section, the organization of drivers for the routine that injects drivers did not accurately
' show the folder structure that should be used. Made a few other changes to the help for clarification purposes.
'
' 18.1.4.163 - September 7, 2021
' Confirmed for sure that there are times where the ei.cfg file is still needed. Fortunately, this was never removed from the code
' as originally planned. We simply gave a user the option to insert an ei.cfg file or not. Removed any comments from the code that
' stated that this file may no longer be needed, but leaving the functionality alone.
'
' 18.1.5.164 - September 8, 2021
' Refined the help message regarding the injection of an EI.CFG file to make the purpose of that file clearer. In addition, in other
' places in the program, when a user wants to see help, any response staring with the letter "H" would be acceptable to get help.
' This did not work for the prompt for an EI.CFG file. This has been changed to be more consistent with the rest of the program.
'
' 18.1.6.165 - September 16, 2021
' Rewrote the "pause" routine. This was a very simple routine that simply printed a blank line and then paused execution of the
' program until the user hit a key. The pause was accomplished by running the command line utility "pause" from a QB64 "shell"
' statement. For most purposes this is fine, except that it would not take input from the QB64 keyboard buffer. As a result,
' taking a series of commands and pasting them into the program would not work when a pause is encountered. This also has the
' potential to cause problems if we want to expand scripting capabilities in the future. The rewrite of this routine solves this.
'
' 18.1.7.166 - September 17, 2021
' We have 2 two "DO...LOOP" structures in the rewritten "PAUSE" code. These loops run forever waiting for a key press and release.
' Revised these structures so that they only look for a press or release 50 times per second to avoid hammering the CPU.
'
' 18.1.8.167 - September 24, 2021
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
' shutdown. Without this, a shutdown cannot be performaed when the system is locked. We also provide a chance to abort shutdown.
'
' 18.2.0.168 - September 26, 2021
' Made some significant changes specifically to better accomodate scripting. If the number of files in a source folder changed,
' this would break scripting because it changes your responses when the program asks if you want to update each individual file.
' To eliminate this problem, the program will now allow you to specify a full path including a file name. This way the program
' does not need to inquire about each file because you are specifying each file unambiguously.
'
' In addition, the script files are now a little easier to read and manually modify.
'
' 18.2.1.169 - September 27, 2021
' Further refinements to the scripting. The script files are much easier to read and edit. Blank lines are now legal and not interpreted
' as an <ENTER>. An <ENTER> is now explicitly shown with an "<ENTER>" tag in the script.
'
' 18.2.2.170 - September 27, 2021
' Bug fix. In the process of adjusting scripting settings, we caused a display problem for those routines that want to show index details
' but do not honor scripting settings (playback, record, or skip). This has been resolved.
'
' 19.0.0.172 - October 5, 2021 (The Windows 11 Release Day Edition)
' This is a major new version!
' The option to create bootable media now has 2 options. The previously existing option to create a single Windows boot option along
' with multiple generic partitions to hold other data remains as was. However, we now have the option to create media that allows for
' multiple bootable partitions. For example, you could boot Windows 10 setup / recovery media, Windows 11 setup / recovery media, Macrium
' reflect recovery media, and any other Windows PE/RE based bootable media, as well as several generic partitions for holding any other
' desired data. This functionality is only for x64 / UEFI based systems and not BIOS or x86 based systems.
'
' 19.0.1.173 - October 5, 2021
' Bug fix. Even if a user did not want to hide operating system and Windows PE / RE partition drive letters, we were hiding them
' anyway. This has been fixed.
'
' 19.0.2.174 - October 11, 2021
' Minor change. Reworded some text in the routine that reorders Windows editions within an image for clarity.
'
' 19.0.3.175 - Octoner 12, 2021
' Fixed a bug where we were skipping over an entire section of code becase we had a call to a subroutine commented out.
'
' 19.0.4.176 - October 12, 2021
' Bug fix: Still dealing with the code to create bootable media. Under a specific set of circumstances we were running the
' code to ask the user what disk write the image to twice. This has been resolved.
'
' 19.0.5.177 - October 13, 2021
' Minor change - Updated some text displayed to user for clarity.
'
' 19.1.0.178 - October 20, 2021
' Microsoft now distributes the the SSU and LCU in one combined package. We have changed our code to work with this new model.
' Note that this may look a bit odd at time becaue we apply the LCU and then apply it again. This is because the first time the
' SSU get applied, the second time, the LCU gets applied. Note that if we were loading certain optional components such as language
' packs, etc., these componts would get installed right between thos two occurences of the LCU update for both WinRE and WinPE (boot.wim).
'
' 19.1.1.179 - October 21, 2021
' Added a couple of lines to the section of code that cleans up project files. In rare circumstances, a user could abort the program
' at a point in time where cleanup of files becomes difficult. We've added a couple lines to try to even more aggressively try to
' cleanup such files.
'
' 19.1.2.180 - October 21, 2021
' Changed some wording of on screen progress indicator. With the new combined SSU and LCU updates, it appears that if an SSU exists,
' the first pass (where the SSU would be applied) completes rather quickly, and then the LCU is added on the second pass which can
' take quite a while. However, if no SSU exists, then it seems that the first pass already applies the LCU and the second pass
' completes almost instantly since there is nothing left to be done. The wording of the progress screens is now changed to reflect
' the fact that the SSU and LCU are combined and we simply call the 2 passes "pass 1" and "pass 2".
'
' 19.1.3.181 - October 23, 2021
' Changed the wording of a few messages in the routine to extract the contents of .CAB files for clarity.
'
' 19.1.4.182 - October 25, 2021
' Changed wording on a menu item.

