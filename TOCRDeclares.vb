'***************************************************************************
' Module:     TOCRDeclares
'
' TOCR declares Version 5.1.0.0

Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices

Module TOCRDeclares

#Region " Structures "

    <StructLayout(LayoutKind.Sequential)>
    Structure TOCRPROCESSOPTIONS_EG
        Dim StructId As Integer
        Dim InvertWholePage As Short
        Dim DeskewOff As Short
        Dim Orientation As Byte
        Dim NoiseRemoveOff As Short
        Dim ReturnNoiseOn As Short
        Dim LineRemoveOff As Short
        Dim DeshadeOff As Short
        Dim InvertOff As Short
        Dim SectioningOn As Short
        Dim MergeBreakOff As Short
        Dim LineRejectOff As Short
        Dim CharacterRejectOff As Short
        Dim ResultsReference As Short
        Dim LexMode As Short
        Dim OCRBOnly As Short
        Dim Speed As Short
        Dim FontStyleInfoOff As Short
        Dim Reserved1 As Short
        Dim Reserved2 As Short
        Dim Reserved3 As Short
        Dim CCAlgorithm As Integer
        Dim CCThreshold As Single
        Dim CGAlgorithm As Integer
        Dim ExtraInfFlags As Integer
        <VBFixedArray(45), MarshalAs(UnmanagedType.ByValArray, SizeConst:=46)>
        Public DisableLangs() As Byte
        Dim Reserved4 As Short
        Dim Reserved5 As Short
        <VBFixedArray(607), MarshalAs(UnmanagedType.ByValArray, SizeConst:=608)>
        Public DisableCharW() As Byte

        Public Sub Initialize()
            ReDim DisableLangs(45)
            ReDim DisableCharW(607)
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
    Public Structure TOCRJOBINFO_EG
        Dim hMMF As IntPtr
        Dim InputFile As String
        Dim StructId As Integer
        Dim JobType As Integer
        Dim PageNo As Integer
        Dim ProcessOptions As TOCRPROCESSOPTIONS_EG

        Public Sub Initialize()
            ProcessOptions.Initialize()
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSHEADER_EG
        Dim StructId As Integer
        Dim XPixelsPerInch As Integer
        Dim YPixelsPerInch As Integer
        Dim NumItems As Integer
        Dim MeanConfidence As Single
        Dim DominantLanguage As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSITEM_EG
        Dim Confidence As Single
        Dim StructId As Short
        Dim OCRCharWUnicode As Short
        Dim OCRCharWInternal As Short
        Dim FontID As Short
        Dim FontStyleInfo As Short
        Dim XPos As Short
        Dim YPos As Short
        Dim XDim As Short
        Dim YDim As Short
        Dim YDimRef As Short
        Dim Noise As Short
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTS_EG
        Dim Hdr As TOCRRESULTSHEADER_EG
        Dim Item() As TOCRRESULTSITEM_EG
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSITEMEXALT_EG
        Dim Factor As Single
        Dim Valid As Short
        Dim OCRCharWUnicode As Short
        Dim OCRCharWInternal As Short
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSITEMEX_EG
        Dim Confidence As Single
        Dim StructId As Short
        Dim OCRCharWUnicode As Short
        Dim OCRCharWInternal As Short
        Dim FontID As Short
        Dim FontStyleInfo As Short
        Dim XPos As Short
        Dim YPos As Short
        Dim XDim As Short
        Dim YDim As Short
        Dim YDimRef As Short
        Dim Noise As Short
        <VBFixedArray(4)> Dim Alt() As TOCRRESULTSITEMEXALT_EG

        Public Sub Initialize()
            ReDim Alt(4)
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSEX_EG
        Dim Hdr As TOCRRESULTSHEADER_EG
        Dim Item() As TOCRRESULTSITEMEX_EG
    End Structure

#End Region

#Region " SUPERSEDED Structures "

    'Superseded by TOCRPROCESSOPTIONS_EG
    <StructLayout(LayoutKind.Sequential)>
    Structure TOCRPROCESSOPTIONS
        Dim StructId As Integer
        Dim InvertWholePage As Short
        Dim DeskewOff As Short
        Dim Orientation As Byte
        Dim NoiseRemoveOff As Short
        Dim LineRemoveOff As Short
        Dim DeshadeOff As Short
        Dim InvertOff As Short
        Dim SectioningOn As Short
        Dim MergeBreakOff As Short
        Dim LineRejectOff As Short
        Dim CharacterRejectOff As Short
        Dim LexOff As Short
        <VBFixedArray(255), MarshalAs(UnmanagedType.ByValArray, SizeConst:=256)>
        Public DisableCharacter() As Short

        Public Sub Initialize()
            ReDim DisableCharacter(255)
        End Sub
    End Structure

    'Superseded by TOCRJOBINFO_EG
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
    Public Structure TOCRJOBINFO2
        Dim StructId As Integer
        Dim JobType As Integer
        Dim InputFile As String
        Dim hMMF As IntPtr
        Dim PageNo As Integer
        Dim ProcessOptions As TOCRPROCESSOPTIONS

        Public Sub Initialize()
            ProcessOptions.Initialize()
        End Sub
    End Structure

    ' Superseded by TOCRJOBINFO2
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
    Public Structure TOCRJOBINFO
        Dim StructId As Integer
        Dim JobType As Integer
        Dim InputFile As String
        Dim PageNo As Integer
        Dim ProcessOptions As TOCRPROCESSOPTIONS

        Public Sub Initialize()
            ProcessOptions.Initialize()
        End Sub
    End Structure

    'Superseded by TOCRRESULTSITEM_EG
    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSITEM
        Dim StructId As Short
        Dim OCRCha As Short
        Dim Confidence As Single
        Dim XPos As Short
        Dim YPos As Short
        Dim XDim As Short
        Dim YDim As Short
    End Structure

    'Superseded by TOCRRESULTSHEADER_EG
    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSHEADER
        Dim StructId As Integer
        Dim XPixelsPerInch As Integer
        Dim YPixelsPerInch As Integer
        Dim NumItems As Integer
        Dim MeanConfidence As Single
    End Structure

    'Superseded by TOCRRESULTS_EG
    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTS
        Dim Hdr As TOCRRESULTSHEADER
        Dim Item() As TOCRRESULTSITEM
    End Structure

    'Superseded by TOCRRESULTSITEMEXALT_EG
    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSITEMEXALT
        Dim Valid As Short
        Dim OCRCha As Short
        Dim Factor As Single
    End Structure

    'Superseded by TOCRRESULTSITEMEX_EG
    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSITEMEX
        Dim StructId As Short
        Dim OCRCha As Short
        Dim Confidence As Single
        Dim XPos As Short
        Dim YPos As Short
        Dim XDim As Short
        Dim YDim As Short
        <VBFixedArray(4)> Dim Alt() As TOCRRESULTSITEMEXALT

        Public Sub Initialize()
            ReDim Alt(4)
        End Sub
    End Structure

    'Superseded by TOCRRESULTSEX_EG
    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSEX
        Dim Hdr As TOCRRESULTSHEADER
        Dim Item() As TOCRRESULTSITEMEX
    End Structure

#End Region

#Region " Declares "

    Declare Function TOCRInitialise Lib "TOCRDll" _
    (ByRef JobNo As Integer) As Integer

    Declare Function TOCRShutdown Lib "TOCRDll" _
        (ByVal JobNo As Integer) As Integer

    Declare Function TOCRDoJob_EG Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobInfo As TOCRJOBINFO_EG) As Integer

    Declare Function TOCRWaitForJob Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobStatus As Integer) As Integer

    Declare Function TOCRWaitForAnyJob Lib "TOCRDll" _
        (ByRef WaitAnyStatus As Integer, ByRef JobNo As Integer) As Integer

    Declare Function TOCRGetJobDBInfo Lib "TOCRDll" _
        (ByVal JobSlotInf As System.IntPtr) As Integer

    Declare Function TOCRGetJobStatus Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobStatus As Integer) As Integer

    Declare Function TOCRGetJobStatusEx Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobStatus As Integer, ByRef Progress As Single, ByRef AutoOrientation As Integer) As Integer

    Declare Ansi Function TOCRGetJobStatusMsg Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Msg As String) As Integer

    Declare Ansi Function TOCRGetNumPages Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Filename As String, ByVal JobType As Integer, ByRef NumPages As Integer) As Integer

    Declare Unicode Function TOCRGetJobResultsEx_EG Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Mode As Integer, ByRef ResultsInf As Integer, ByVal Bytes As System.IntPtr) As Integer

    Declare Ansi Function TOCRGetLicenceInfoEx Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Licence As String, ByRef Volume As Integer, ByRef Time As Integer, ByRef Remaining As Integer, ByRef Features As Integer) As Integer

    ' Convert a TIF or PDF file to a bitmap file
    Declare Ansi Function TOCRConvertFormat Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal InputAddr As String, ByVal InputFormat As Integer, ByVal OutputAddr As String, ByVal OutputFormat As Integer, ByVal PageNo As Integer) As Integer
    ' Convert a TIF or PDF file to a memory mapped file handle
    Declare Ansi Function TOCRConvertFormat Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal InputAddr As String, ByVal InputFormat As Integer, ByRef OutputAddr As System.IntPtr, ByVal OutputFormat As Integer, ByVal PageNo As Integer) As Integer

    ' These functions cannot be used to get/set the log file name in x64
    Declare Function TOCRSetConfig Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Parameter As Integer, ByVal Value As Integer) As Integer
    Declare Function TOCRGetConfig Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Parameter As Integer, ByRef Value As Integer) As Integer

    ' These functions can be used to get/set the log file name in x64
    Declare Ansi Function TOCRSetConfigStr Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Parameter As Integer, ByVal Value As String) As Integer
    Declare Ansi Function TOCRGetConfigStr Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Parameter As Integer, ByVal Value As String) As Integer

    Declare Function TOCRTWAINAcquire Lib "TOCRDll" _
        (ByRef NumberOfImages As Integer) As Integer

    Declare Function TOCRTWAINGetImages Lib "TOCRDll" _
        (ByVal GlobalMemoryDIBs As System.IntPtr) As Integer

    Declare Function TOCRTWAINSelectDS Lib "TOCRDll" _
        () As Integer

    Declare Function TOCRTWAINShowUI Lib "TOCRDll" _
        (ByVal Show As Short) As Integer

#End Region

#Region " SUPERSEDED Declares "

    ' Superseded by TOCRDoJob_EG
    Declare Function TOCRDoJob2 Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobInfo As TOCRJOBINFO2) As Integer

    ' Superseded by TOCRDoJob2
    Declare Function TOCRDoJob Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobInfo As TOCRJOBINFO) As Integer

    ' Superseded by TOCRGetJobResultsEx_EG Mode
    Declare Function TOCRGetJobResults Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef ResultsInf As Integer, ByVal Bytes As System.IntPtr) As Integer

    ' Superseded by TOCRGetJobResultsEx_EG
    Declare Function TOCRGetJobResultsEx Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Mode As Integer, ByRef ResultsInf As Integer, ByVal Bytes As System.IntPtr) As Integer

    ' Superseded by TOCRGetConfig
    Declare Function TOCRGetErrorMode Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef ErrorMode As Integer) As Integer

    ' Superseded by TOCRSetConfig
    Declare Function TOCRSetErrorMode Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal ErrorMode As Integer) As Integer

    ' UNTESTED REDUNDANT - use the Bitmap class in .NET
    'Declare Function TOCRRotateMonoBitmap Lib "TOCRDll" _
    '    (ByRef hBmp As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal Orientation As Integer) As Integer

    ' UNTESTED - obsolete, use TOCRConvertFormat
    'Declare Ansi Function TOCRConvertTIFFtoDIB Lib "TOCRDll" _
    '    (ByVal JobNo As Integer, ByVal InputFilename As String, ByVal OutputFilename As String, ByVal PageNo As Integer) As Integer

    ' UNTESTED - Superseded by TOCRGetLicenceInfoEx
    'Declare Function TOCRGetLicenceInfo Lib "TOCRDll" _
    '    (ByRef NumOfJobSlots As Integer, ByRef Volume As Integer, ByRef Time As Integer, ByRef Remaining As Integer) As Integer

#End Region

#Region " User constants "
    Public Const TOCRJOBMSGLENGTH As Short = 512        ' max length of a job status message
    Public Const TOCRFONTNAMELENGTH As Integer = 65     ' max length of a returned font name

    Public Const TOCRMAXPPM As Integer = 78741          ' max pixels per metre
    Public Const TOCRMINPPM As Integer = 984            ' min pixels per metre

    ' Setting for JobNo for TOCRSetErrorMode and TOCRGetErrorMode
    Public Const TOCRDEFERRORMODE As Integer = -1       ' set/get the API error mode for all jobs

    ' Settings for ErrorMode for TOCRSetErrorMode and TOCRGetErrorMode
    Public Const TOCRERRORMODE_NONE As Integer = 0      ' API errors unseen (use return status of API calls)
    Public Const TOCRERRORMODE_MSGBOX As Integer = 1    ' API errors will bring up a message box
    Public Const TOCRERRORMODE_LOG As Integer = 2       ' errors are sent to a log file

    ' Setting for TOCRShutdown
    Public Const TOCRSHUTDOWNALL As Integer = -1        ' stop and shutdown processing for all jobs

    ' Values returned by TOCRGetJobStatus JobStatus
    Public Const TOCRJOBSTATUS_ERROR As Integer = -1    ' an error ocurred
    Public Const TOCRJOBSTATUS_BUSY As Integer = 0      ' the job is still processing
    Public Const TOCRJOBSTATUS_DONE As Integer = 1      ' the job completed successfully
    Public Const TOCRJOBSTATUS_IDLE As Integer = 2      ' no job has been specified yet

    ' Settings for TOCRJOBINFO.JobType
    Public Const TOCRJOBTYPE_TIFFFILE As Integer = 0    ' TOCRJOBINFO.InputFile specifies a tiff file
    Public Const TOCRJOBTYPE_DIBFILE As Integer = 1     ' TOCRJOBINFO.InputFile specifies a dib (bmp) file
    Public Const TOCRJOBTYPE_DIBCLIPBOARD As Integer = 2 ' clipboard contains a dib (clipboard format CF_DIB)
    Public Const TOCRJOBTYPE_MMFILEHANDLE As Integer = 3 ' TOCRJOBINFO.PageNo specifies a handle to a memory mapped DIB file
    Public Const TOCRJOBTYPE_PDFFILE As Integer = 4    ' TOCRJOBINFO.InputFile specifies a PDF file

    ' Settings for TOCRJOBINFO.PROCESSOPTIONS.Orientation
    ' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.Orientation
    Public Const TOCRJOBORIENT_AUTO As Byte = 0         ' detect orientation and rotate automatically
    Public Const TOCRJOBORIENT_OFF As Byte = 255        ' don't rotate
    Public Const TOCRJOBORIENT_90 As Byte = 1           ' 90 degrees clockwise rotation
    Public Const TOCRJOBORIENT_180 As Byte = 2          ' 180 degrees clockwise rotation
    Public Const TOCRJOBORIENT_270 As Byte = 3          ' 270 degrees clockwise rotation

    ' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.ResultsReference
    Public Const TOCRRESULTSREFERENCE_SELFREL As Byte = 0 ' relative to the first top left character recognised
    Public Const TOCRRESULTSREFERENCE_BEFORE As Byte = 1  ' page position before rotation and deskewing
    Public Const TOCRRESULTSREFERENCE_BETWEEN As Byte = 2 ' page position after rotation but before deskewing deskewing
    Public Const TOCRRESULTSREFERENCE_AFTER As Byte = 3   ' page position after rotation and deskewing

    ' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.LexMode
    Public Const TOCRJOBLEXMODE_AUTO As Byte = 0          ' decide whether to apply lex
    'Public Const TOCRJOBLEXMODE_ON As Byte = 1            ' lex always on - removed for v5
    'Public Const TOCRJOBLEXMODE_OFF As Byte = 2           ' lex always off - removed for v5

    ' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.Speed
    Public Const TOCRJOBSPEED_SLOW As Byte = 0
    Public Const TOCRJOBSPEED_MEDIUM As Byte = 1
    Public Const TOCRJOBSPEED_FAST As Byte = 2
    Public Const TOCRJOBSPEED_EXPRESS As Byte = 3

    ' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.CCAlgorithm (Colour Conversion Algorithm)
    Public Const TOCRJOBCC_AVERAGE As Integer = 0           ' (R+G+3)/3
    Public Const TOCRJOBCC_LUMA_BT601 As Integer = 1        ' 0.299*R + 0.587*G + 0.114*B
    Public Const TOCRJOBCC_LUMA_BT709 As Integer = 2        ' 0.2126*R + 0.7152*G + 0.0722*B
    Public Const TOCRJOBCC_DESATURATION As Integer = 3      ' (max(R,G,B) + min(R,G,B))/2
    Public Const TOCRJOBCC_DECOMPOSITION_MAX As Integer = 4 ' max(R,G,B)
    Public Const TOCRJOBCC_DECOMPOSITION_MIN As Integer = 5 ' min(R,G,B)
    Public Const TOCRJOBCC_RED As Integer = 6               ' R
    Public Const TOCRJOBCC_GREEN As Integer = 7             ' G
    Public Const TOCRJOBCC_BLUE As Integer = 8              ' B

    ' Settings for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.CGAlgorithm (Greyscale Conversion Algorithm)
    Public Const TOCRJOBCG_HISTOGRAM As Integer = 9
    Public Const TOCRJOBCG_REGIONS As Integer = 10

    ' Flags for TOCRJOBINFO_EG.PROCESSOPTIONS_EG.ExtraInfFlags
    Public Const TOCREXTRAINF_RETURNBITMAP1 As Byte = 1

    ' Values returned by TOCRGetJobDBInfo
    Public Const TOCRJOBSLOT_FREE As Integer = 0        ' job slot is free for use
    Public Const TOCRJOBSLOT_OWNEDBYYOU As Integer = 1  ' job slot is in use by your process
    Public Const TOCRJOBSLOT_BLOCKEDBYYOU As Integer = 2 ' blocked by own process (re-initialise)
    Public Const TOCRJOBSLOT_OWNEDBYOTHER As Integer = -1 ' job slot is in use by another process (can't use)
    Public Const TOCRJOBSLOT_BLOCKEDBYOTHER As Integer = -2 ' blocked by another process (can't use)

    ' Values returned in WaitAnyStatus by TOCRWaitForAnyJob
    Public Const TOCRWAIT_OK As Integer = 0             ' JobNo is the job that finished (get and check it's JobStatus)
    Public Const TOCRWAIT_SERVICEABORT As Integer = 1   ' JobNo is the job that failed (re-initialise)
    Public Const TOCRWAIT_CONNECTIONBROKEN As Integer = 2 ' JobNo is the job that failed (re-initialise)
    Public Const TOCRWAIT_FAILED As Integer = -1        ' JobNo not set - check manually
    Public Const TOCRWAIT_NOJOBSFOUND As Integer = -2   ' JobNo not set - no running jobs found

    ' Settings for Mode for TOCRGetJobResultsEx
    Public Const TOCRGETRESULTS_NORMAL As Integer = 0   ' return results for TOCRRESULTS
    Public Const TOCRGETRESULTS_EXTENDED As Integer = 1 ' return results for TOCRRESULTSEX

    ' Settings for Mode for TOCRGetJobResultsEx_EG
    Public Const TOCRGETRESULTS_NORMAL_EG As Integer = 2   ' return results for TOCRRESULTS_EG
    Public Const TOCRGETRESULTS_EXTENDED_EG As Integer = 3 ' return results for TOCRRESULTSEX_EG

    ' Values returned in ResultsInf by TOCRGetJobResults and TOCRGetJobResultsEx
    Public Const TOCRGETRESULTS_NORESULTS As Integer = -1 ' no results are available

    ' Flags returned by TOCRResults_EG.Item().FontStyleInfo
    ' Flags returned by TOCRResultsEx_EG.Item().FontStyleInfo
    Public Const TOCRRESULTSFONT_NOTSET As UShort = 0   ' character tyle is not specified
    Public Const TOCRRESULTSFONT_NORMAL As UShort = 1   ' character is Normal
    Public Const TOCRRESULTSFONT_ITALIC As UShort = 2   ' character is Italic
    Public Const TOCRRESULTSFONT_UNDERLINE As UShort = 4 ' character is Underlined

    ' Values for TOCRConvertFormat InputFormat
    Public Const TOCRCONVERTFORMAT_TIFFFILE As Integer = TOCRJOBTYPE_TIFFFILE
    Public Const TOCRCONVERTFORMAT_PDFFILE As Integer = TOCRJOBTYPE_PDFFILE

    ' Values for TOCRConvertFormat OutputFormat
    Public Const TOCRCONVERTFORMAT_DIBFILE As Integer = TOCRJOBTYPE_DIBFILE
    Public Const TOCRCONVERTFORMAT_MMFILEHANDLE As Integer = TOCRJOBTYPE_MMFILEHANDLE

    ' Values for licence features (returned by TOCRGetLicenceInfoEx)
    Public Const TOCRLICENCE_STANDARD As Integer = 1    ' standard licence (no higher characters)
    Public Const TOCRLICENCE_EURO As Integer = 2        ' higher characters
    Public Const TOCRLICENCE_EUROUPGRADE As Integer = 3 ' standard licence upgraded to euro
    Public Const TOCRLICENCE_V3SE As Integer = 4        ' V3SE version 3 standard edition licence (no API)
    Public Const TOCRLICENCE_V3SEUPGRADE As Integer = 5 ' versions 1/2 upgraded to V3 standard edition (no API)
    ' Note V4 licences are the same as V3 Pro licences
    Public Const TOCRLICENCE_V3PRO As Integer = 6       ' V3PRO version 3 pro licence
    Public Const TOCRLICENCE_V3PROUPGRADE As Integer = 7 ' versions 1/2 upgraded to version 3 pro
    Public Const TOCRLICENCE_V3SEPROUPGRADE As Integer = 8 ' version 3 standard edition upgraded to version 3 pro
    Public Const TOCRLICENCE_V5 As Integer = 9           ' version 5
    Public Const TOCRLICENCE_V5UPGRADE3 As Integer = 10  ' version 5 upgraded from version 3
    Public Const TOCRLICENCE_V5UPGRADE12 As Integer = 11 ' version 5 upgraded from version 1/2

    ' Values for TOCRSetConfig and TOCRGetConfig
    Public Const TOCRCONFIG_DEFAULTJOB As Integer = -1  ' default job number (all new jobs)
    Public Const TOCRCONFIG_DLL_ERRORMODE As Integer = 0 ' set the dll ErrorMode
    Public Const TOCRCONFIG_SRV_ERRORMODE As Integer = 1 ' set the service ErrorMode
    Public Const TOCRCONFIG_SRV_THREADPRIORITY As Integer = 2 ' set the service thread priority
    Public Const TOCRCONFIG_DLL_MUTEXWAIT As Integer = 3 ' set the dll mutex wait timeout (ms)
    Public Const TOCRCONFIG_DLL_EVENTWAIT As Integer = 4 ' set the dll event wait timeout (ms)
    Public Const TOCRCONFIG_SRV_MUTEXWAIT As Integer = 5 ' set the service mutex wait timeout (ms)
    Public Const TOCRCONFIG_LOGFILE As Integer = 6      ' set the log file name
#End Region

#Region " Error Codes "
    Public Const TOCR_OK As Integer = 0

    ' Error codes returned by an API function
    'Public Const TOCRERR_ILLEGALJOBNO As Integer = 1
    'Public Const TOCRERR_FAILLOCKDB As Integer = 2
    'Public Const TOCRERR_NOFREEJOBSLOTS As Integer = 3
    'Public Const TOCRERR_FAILSTARTSERVICE As Integer = 4
    'Public Const TOCRERR_FAILINITSERVICE As Integer = 5
    'Public Const TOCRERR_JOBSLOTNOTINIT As Integer = 6
    'Public Const TOCRERR_JOBSLOTINUSE As Integer = 7
    'Public Const TOCRERR_SERVICEABORT As Integer = 8
    'Public Const TOCRERR_CONNECTIONBROKEN As Integer = 9
    'Public Const TOCRERR_INVALIDSTRUCTID As Integer = 10
    'Public Const TOCRERR_FAILGETVERSION As Integer = 11
    'Public Const TOCRERR_FAILLICENCEINF As Integer = 12
    'Public Const TOCRERR_LICENCEEXCEEDED As Integer = 13
    'Public Const TOCRERR_MISMATCH As Integer = 15
    'Public Const TOCRERR_JOBSLOTNOTYOURS As Integer = 16

    'Public Const TOCRERR_FAILGETJOBSTATUS1 As Integer = 20
    'Public Const TOCRERR_FAILGETJOBSTATUS2 As Integer = 21
    'Public Const TOCRERR_FAILGETJOBSTATUS3 As Integer = 22
    'Public Const TOCRERR_FAILCONVERT As Integer = 23
    'Public Const TOCRERR_FAILSETCONFIG As Integer = 24
    'Public Const TOCRERR_FAILGETCONFIG As Integer = 25

    'Public Const TOCRERR_FAILDOJOB1 As Integer = 30
    'Public Const TOCRERR_FAILDOJOB2 As Integer = 31
    'Public Const TOCRERR_FAILDOJOB3 As Integer = 32
    'Public Const TOCRERR_FAILDOJOB4 As Integer = 33
    'Public Const TOCRERR_FAILDOJOB5 As Integer = 34
    'Public Const TOCRERR_FAILDOJOB6 As Integer = 35
    'Public Const TOCRERR_FAILDOJOB7 As Integer = 36
    'Public Const TOCRERR_FAILDOJOB8 As Integer = 37
    'Public Const TOCRERR_FAILDOJOB9 As Integer = 38
    'Public Const TOCRERR_FAILDOJOB10 As Integer = 39
    'Public Const TOCRERR_UNKNOWNJOBTYPE1 As Integer = 40
    'Public Const TOCRERR_JOBNOTSTARTED1 As Integer = 41
    'Public Const TOCRERR_FAILDUPHANDLE As Integer = 42

    'Public Const TOCRERR_FAILGETJOBSTATUSMSG1 As Integer = 45
    'Public Const TOCRERR_FAILGETJOBSTATUSMSG2 As Integer = 46

    'Public Const TOCRERR_FAILGETNUMPAGES1 As Integer = 50
    'Public Const TOCRERR_FAILGETNUMPAGES2 As Integer = 51
    'Public Const TOCRERR_FAILGETNUMPAGES3 As Integer = 52
    'Public Const TOCRERR_FAILGETNUMPAGES4 As Integer = 53
    'Public Const TOCRERR_FAILGETNUMPAGES5 As Integer = 54

    'Public Const TOCRERR_FAILGETRESULTS1 As Integer = 60
    'Public Const TOCRERR_FAILGETRESULTS2 As Integer = 61
    'Public Const TOCRERR_FAILGETRESULTS3 As Integer = 62
    'Public Const TOCRERR_FAILGETRESULTS4 As Integer = 63
    'Public Const TOCRERR_FAILALLOCMEM100 As Integer = 64
    'Public Const TOCRERR_FAILALLOCMEM101 As Integer = 65
    'Public Const TOCRERR_FILENOTSPECIFIED As Integer = 66
    'Public Const TOCRERR_INPUTNOTSPECIFIED As Integer = 67
    'Public Const TOCRERR_OUTPUTNOTSPECIFIED As Integer = 68
    'Public Const TOCRERR_INVALIDPARAMETER As Integer = 69

    'Public Const TOCRERR_FAILROTATEBITMAP As Integer = 70

    'Public Const TOCERR_TWAINPARTIALACQUIRE As Integer = 80
    'Public Const TOCERR_TWAINFAILEDACQUIRE As Integer = 81
    'Public Const TOCERR_TWAINNOIMAGES As Integer = 82
    'Public Const TOCERR_TWAINSELECTDSFAILED As Integer = 83
    'Public Const TOCERR_MMFNOTALLOWED As Integer = 84
    'Public Const TOCRERR_ILLEGALFONTID As Integer = 85

    'Public Const TOCRERR_FAILGETMMF As Integer = 90
    'Public Const TOCRERR_MMFNOTAVAILABLE As Integer = 91

    'Public Const TOCRERR_PDFEXTRACTOR As Integer = 95
    'Public Const TOCRERR_PDFERROR2 As Integer = 96
    'Public Const TOCRERR_PDFARCHIVER As Integer = 97


    'Public Const TOCRERR_FONTSNOTLOADED As Integer = -2

    '' Error codes which may be seen in a msgbox or console but will not be returned by an API function
    'Public Const TOCRERR_INVALIDSERVICESTART As Integer = 1000
    'Public Const TOCRERR_FAILSERVICEINIT As Integer = 1001
    'Public Const TOCRERR_FAILLICENCE1 As Integer = 1002
    'Public Const TOCRERR_FAILSERVICESTART As Integer = 1003
    'Public Const TOCRERR_UNKNOWNCMD As Integer = 1004
    'Public Const TOCRERR_FAILREADCOMMAND As Integer = 1005
    'Public Const TOCRERR_FAILREADOPTIONS As Integer = 1006
    'Public Const TOCRERR_FAILWRITEJOBSTATUS1 As Integer = 1007
    'Public Const TOCRERR_FAILWRITEJOBSTATUS2 As Integer = 1008
    'Public Const TOCRERR_FAILWRITETHREADH As Integer = 1009
    'Public Const TOCRERR_FAILREADJOBINFO1 As Integer = 1010
    'Public Const TOCRERR_FAILREADJOBINFO2 As Integer = 1011
    'Public Const TOCRERR_FAILREADJOBINFO3 As Integer = 1012
    'Public Const TOCRERR_FAILWRITEPROGRESS As Integer = 1013
    'Public Const TOCRERR_FAILWRITEJOBSTATUSMSG As Integer = 1014
    'Public Const TOCRERR_FAILWRITERESULTSSIZE As Integer = 1015
    'Public Const TOCRERR_FAILWRITERESULTS As Integer = 1016
    'Public Const TOCRERR_FAILWRITEAUTOORIENT As Integer = 1017
    'Public Const TOCRERR_FAILLICENCE2 As Integer = 1018
    'Public Const TOCRERR_FAILLICENCE3 As Integer = 1019

    'Public Const TOCRERR_TOOMANYCOLUMNS As Integer = 1020
    'Public Const TOCRERR_TOOMANYROWS As Integer = 1021
    'Public Const TOCRERR_EXCEEDEDMAXZONE As Integer = 1022
    'Public Const TOCRERR_NSTACKTOOSMALL As Integer = 1023
    'Public Const TOCRERR_ALGOERR1 As Integer = 1024
    'Public Const TOCRERR_ALGOERR2 As Integer = 1025
    'Public Const TOCRERR_EXCEEDEDMAXCP As Integer = 1026
    'Public Const TOCRERR_CANTFINDPAGE As Integer = 1027
    'Public Const TOCRERR_UNSUPPORTEDIMAGETYPE As Integer = 1028
    'Public Const TOCRERR_IMAGETOOWIDE As Integer = 1029
    'Public Const TOCRERR_IMAGETOOLONG As Integer = 1030
    'Public Const TOCRERR_UNKNOWNJOBTYPE2 As Integer = 1031
    'Public Const TOCRERR_TOOWIDETOROT As Integer = 1032
    'Public Const TOCRERR_TOOLONGTOROT As Integer = 1033
    'Public Const TOCRERR_INVALIDPAGENO As Integer = 1034
    'Public Const TOCRERR_FAILREADJOBTYPENUMBYTES As Integer = 1035
    'Public Const TOCRERR_FAILREADFILENAME As Integer = 1036
    'Public Const TOCRERR_FAILSENDNUMPAGES As Integer = 1037
    'Public Const TOCRERR_FAILOPENCLIP As Integer = 1038
    'Public Const TOCRERR_NODIBONCLIP As Integer = 1039
    'Public Const TOCRERR_FAILREADDIBCLIP As Integer = 1040
    'Public Const TOCRERR_FAILLOCKDIBCLIP As Integer = 1041
    'Public Const TOCRERR_UNKOWNDIBFORMAT As Integer = 1042
    'Public Const TOCRERR_FAILREADDIB As Integer = 1043
    'Public Const TOCRERR_NOXYPPM As Integer = 1044
    'Public Const TOCRERR_FAILCREATEDIB As Integer = 1045
    'Public Const TOCRERR_FAILWRITEDIBCLIP As Integer = 1046
    'Public Const TOCRERR_FAILALLOCMEMDIB As Integer = 1047
    'Public Const TOCRERR_FAILLOCKMEMDIB As Integer = 1048
    'Public Const TOCRERR_FAILCREATEFILE As Integer = 1049
    'Public Const TOCRERR_FAILOPENFILE1 As Integer = 1050
    'Public Const TOCRERR_FAILOPENFILE2 As Integer = 1051
    'Public Const TOCRERR_FAILOPENFILE3 As Integer = 1052
    'Public Const TOCRERR_FAILOPENFILE4 As Integer = 1053
    'Public Const TOCRERR_FAILREADFILE1 As Integer = 1054
    'Public Const TOCRERR_FAILREADFILE2 As Integer = 1055
    'Public Const TOCRERR_FAILFINDDATA1 As Integer = 1056
    'Public Const TOCRERR_TIFFERROR1 As Integer = 1057
    'Public Const TOCRERR_TIFFERROR2 As Integer = 1058
    'Public Const TOCRERR_TIFFERROR3 As Integer = 1059
    'Public Const TOCRERR_TIFFERROR4 As Integer = 1060
    'Public Const TOCRERR_FAILREADDIBHANDLE As Integer = 1061
    'Public Const TOCRERR_PAGETOOBIG As Integer = 1062
    'Public Const TOCRERR_FAILSETTHREADPRIORITY As Integer = 1063
    'Public Const TOCRERR_FAILSETSRVERRORMODE As Integer = 1064
    'Public Const TOCRERR_FAILSENDFONT1 As Integer = 1065
    'Public Const TOCRERR_FAILSENDFONT2 As Integer = 1066
    'Public Const TOCRERR_FAILSENDFONT3 As Integer = 1067
    'Public Const TOCRERR_FAILALLOCFONTMEM As Integer = 1068
    'Public Const TOCRERR_FAILWRITEEXTRAINF As Integer = 1069

    'Public Const TOCRERR_FAILREADFILENAME1 As Integer = 1070
    'Public Const TOCRERR_FAILREADFILENAME2 As Integer = 1071
    'Public Const TOCRERR_FAILREADFILENAME3 As Integer = 1072
    'Public Const TOCRERR_FAILREADFILENAME4 As Integer = 1073
    'Public Const TOCRERR_FAILREADFILENAME5 As Integer = 1074

    'Public Const TOCRERR_FAILREADFORMAT1 As Integer = 1080
    'Public Const TOCRERR_FAILREADFORMAT2 As Integer = 1081

    'Public Const TOCRERR_FAILALLOCMEM1 As Integer = 1101
    'Public Const TOCRERR_FAILALLOCMEM2 As Integer = 1102
    'Public Const TOCRERR_FAILALLOCMEM3 As Integer = 1103
    'Public Const TOCRERR_FAILALLOCMEM4 As Integer = 1104
    'Public Const TOCRERR_FAILALLOCMEM5 As Integer = 1105
    'Public Const TOCRERR_FAILALLOCMEM6 As Integer = 1106
    'Public Const TOCRERR_FAILALLOCMEM7 As Integer = 1107
    'Public Const TOCRERR_FAILALLOCMEM8 As Integer = 1108
    'Public Const TOCRERR_FAILALLOCMEM9 As Integer = 1109
    'Public Const TOCRERR_FAILALLOCMEM10 As Integer = 1110

    'Public Const TOCRERR_FAILWRITEMMFH As Integer = 1150
    'Public Const TOCRERR_FAILREADACK As Integer = 1151
    'Public Const TOCRERR_FAILFILEMAP As Integer = 1152
    'Public Const TOCRERR_FAILFILEVIEW As Integer = 1153
    'Public Const TOCRERR_FAILSENDBMP As Integer = 1154

    'Public Const TOCRERR_FAILREADFILE3 As Integer = 1155
    'Public Const TOCRERR_FAILREADFILE4 As Integer = 1156

    'Public Const TOCRERR_PDFERROR1 As Integer = 1157

    'Public Const TOCRERR_BUFFEROVERFLOW1 As Integer = 2001

    'Public Const TOCRERR_MAPOVERFLOW As Integer = 2002
    'Public Const TOCRERR_REBREAKNEXTCALL As Integer = 2003
    'Public Const TOCRERR_REBREAKNEXTDATA As Integer = 2004
    'Public Const TOCRERR_REBREAKEXACTCALL As Integer = 2005
    'Public Const TOCRERR_MAXZCANOVERFLOW1 As Integer = 2006
    'Public Const TOCRERR_MAXZCANOVERFLOW2 As Integer = 2007
    'Public Const TOCRERR_BUFFEROVERFLOW2 As Integer = 2008
    'Public Const TOCRERR_NUMKCOVERFLOW As Integer = 2009
    'Public Const TOCRERR_BUFFEROVERFLOW3 As Integer = 2010
    'Public Const TOCRERR_BUFFEROVERFLOW4 As Integer = 2011
    'Public Const TOCRERR_SEEDERROR As Integer = 2012

    'Public Const TOCRERR_FCZYREF As Integer = 2020
    'Public Const TOCRERR_MAXTEXTLINES1 As Integer = 2021
    'Public Const TOCRERR_LINEINDEX As Integer = 2022
    'Public Const TOCRERR_MAXFCZSONLINE As Integer = 2023
    'Public Const TOCRERR_MEMALLOC1 As Integer = 2024
    'Public Const TOCRERR_MERGEBREAK As Integer = 2025

    'Public Const TOCRERR_DKERNPRANGE1 As Integer = 2030
    'Public Const TOCRERR_DKERNPRANGE2 As Integer = 2031
    'Public Const TOCRERR_BUFFEROVERFLOW5 As Integer = 2032
    'Public Const TOCRERR_BUFFEROVERFLOW6 As Integer = 2033

    'Public Const TOCRERR_FILEOPEN1 As Integer = 2040
    'Public Const TOCRERR_FILEOPEN2 As Integer = 2041
    'Public Const TOCRERR_FILEOPEN3 As Integer = 2042
    'Public Const TOCRERR_FILEREAD1 As Integer = 2043
    'Public Const TOCRERR_FILEREAD2 As Integer = 2044
    'Public Const TOCRERR_SPWIDZERO As Integer = 2045
    'Public Const TOCRERR_FAILALLOCMEMLEX1 As Integer = 2046
    'Public Const TOCRERR_FAILALLOCMEMLEX2 As Integer = 2047

    'Public Const TOCRERR_BADOBWIDTH As Integer = 2050
    'Public Const TOCRERR_BADROTATION As Integer = 2051

    'Public Const TOCRERR_REJHIDMEMALLOC As Integer = 2055

    'Public Const TOCRERR_UIDA As Integer = 2070
    'Public Const TOCRERR_UIDB As Integer = 2071
    'Public Const TOCRERR_ZEROUID As Integer = 2072
    'Public Const TOCRERR_CERTAINTYDBNOTINIT As Integer = 2073
    'Public Const TOCRERR_MEMALLOCINDEX As Integer = 2074
    'Public Const TOCRERR_CERTAINTYDB_INIT As Integer = 2075
    'Public Const TOCRERR_CERTAINTYDB_DELETE As Integer = 2076
    'Public Const TOCRERR_CERTAINTYDB_INSERT1 As Integer = 2077
    'Public Const TOCRERR_CERTAINTYDB_INSERT2 As Integer = 2078
    'Public Const TOCRERR_OPENXORNEAREST As Integer = 2079
    'Public Const TOCRERR_XORNEAREST As Integer = 2079

    'Public Const TOCRERR_OPENSETTINGS As Integer = 2080
    'Public Const TOCRERR_READSETTINGS1 As Integer = 2081
    'Public Const TOCRERR_READSETTINGS2 As Integer = 2082
    'Public Const TOCRERR_BADSETTINGS As Integer = 2083
    'Public Const TOCRERR_WRITESETTINGS As Integer = 2084
    'Public Const TOCRERR_MAXSCOREDIFF As Integer = 2085

    'Public Const TOCRERR_YDIMREFZERO1 As Integer = 2090
    'Public Const TOCRERR_YDIMREFZERO2 As Integer = 2091
    'Public Const TOCRERR_YDIMREFZERO3 As Integer = 2092
    'Public Const TOCRERR_ASMFILEOPEN As Integer = 2093
    'Public Const TOCRERR_ASMFILEREAD As Integer = 2094
    'Public Const TOCRERR_MEMALLOCASM As Integer = 2095
    'Public Const TOCRERR_MEMREALLOCASM As Integer = 2096
    'Public Const TOCRERR_SDBFILEOPEN As Integer = 2097
    'Public Const TOCRERR_SDBFILEREAD As Integer = 2098
    'Public Const TOCRERR_SDBFILEBAD1 As Integer = 2099
    'Public Const TOCRERR_SDBFILEBAD2 As Integer = 2100
    'Public Const TOCRERR_MEMALLOCSDB As Integer = 2101
    'Public Const TOCRERR_DEVEL1 As Integer = 2102
    'Public Const TOCRERR_DEVEL2 As Integer = 2103
    'Public Const TOCRERR_DEVEL3 As Integer = 2104
    'Public Const TOCRERR_DEVEL4 As Integer = 2105
    'Public Const TOCRERR_DEVEL5 As Integer = 2106
    'Public Const TOCRERR_DEVEL6 As Integer = 2107
    'Public Const TOCRERR_DEVEL7 As Integer = 2108
    'Public Const TOCRERR_DEVEL8 As Integer = 2109
    'Public Const TOCRERR_DEVEL9 As Integer = 2110
    'Public Const TOCRERR_DEVEL10 As Integer = 2111
    'Public Const TOCRERR_DEVEL11 As Integer = 2112
    'Public Const TOCRERR_DEVEL12 As Integer = 2113
    'Public Const TOCRERR_DEVEL13 As Integer = 2114
    'Public Const TOCRERR_FILEOPEN4 As Integer = 2115
    'Public Const TOCRERR_FILEOPEN5 As Integer = 2116
    'Public Const TOCRERR_FILEOPEN6 As Integer = 2117
    'Public Const TOCRERR_FILEREAD3 As Integer = 2118
    'Public Const TOCRERR_FILEREAD4 As Integer = 2119
    'Public Const TOCRERR_ZOOMGTOOBIG As Integer = 2120
    'Public Const TOCRERR_ZOOMGOUTOFRANGE As Integer = 2121

    'Public Const TOCRERR_MEMALLOCRESULTS As Integer = 2130

    'Public Const TOCRERR_MEMALLOCHEAP As Integer = 2140
    'Public Const TOCRERR_HEAPNOTINITIALISED As Integer = 2141
    'Public Const TOCRERR_MEMLIMITHEAP As Integer = 2142
    'Public Const TOCRERR_MEMREALLOCHEAP As Integer = 2143
    'Public Const TOCRERR_MEMALLOCFCZBM As Integer = 2144
    'Public Const TOCRERR_FCZBMOVERLAP As Integer = 2145
    'Public Const TOCRERR_FCZBMLOCATION As Integer = 2146
    'Public Const TOCRERR_MEMREALLOCFCZBM As Integer = 2147
    'Public Const TOCRERR_MEMALLOCFCHBM As Integer = 2148
    'Public Const TOCRERR_MEMREALLOCFCHBM As Integer = 2149

#End Region

End Module
