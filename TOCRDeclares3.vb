'***************************************************************************
' Module:     TOCRDeclares
'
' TOCR declares Version 5.0.0.0

#Const SUPERSEDED = False ' disallow superseded routines

Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices

Module TOCRDeclares

#Region " Structures "
    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRPROCESSPDFOPTIONS_EG
        Dim ResultsOn As Byte 'V5 addition
        Dim OriginalImageOn As Byte 'V5 addition
        Dim ProcessedImageOn As Byte 'V5 addition
        Dim PDFSpare As Integer 'V5 addition
    End Structure


    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRCHAROPTIONS_EG
        'Line below: V5 extended from 601 to 608
        <VBFixedArray(607), MarshalAs(UnmanagedType.ByValArray, SizeConst:=608)>
        Public DisableCharW() As Byte

        Public Sub Initialize()
            ReDim DisableCharW(607) 'V5 extended 600 to 607
        End Sub
    End Structure 'TOCRCHAROPTIONS_EG

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRLANGUAGEOPTIONS_EG
        'Line below: V5 extended from 601 to 608
        '<VBFixedArray(607), MarshalAs(UnmanagedType.ByValArray, SizeConst:=608)> _
        'Public DisableCharW() As Byte
        ' 608 bytes
        <VBFixedArray(45), MarshalAs(UnmanagedType.ByValArray, SizeConst:=46)>
        Dim DisableLangs() As Byte
        ' 654 bytes

        Public Sub Initialize()
            ReDim DisableLangs(45)
            'ReDim DisableCharW(607) 'V5 extended 600 to 607
        End Sub
    End Structure 'TOCRLANGUAGEOPTIONS_EG

    <StructLayout(LayoutKind.Sequential)>
    Structure TOCRPROCESSOPTIONS_EG
        Dim StructId As Int32
        ' 4 bytes
        Dim InvertWholePage As Byte
        Dim DeskewOff As Byte
        Dim Orientation As Byte
        Dim NoiseRemoveOff As Byte
        ' 8 bytes
        Dim ReturnNoiseOn As Byte 'v5 addition
        Dim LineRemoveOff As Byte
        Dim DeshadeOff As Byte
        Dim InvertOff As Byte
        ' 12 bytes
        Dim SectioningOn As Byte
        Dim MergeBreakOff As Byte
        Dim LineRejectOff As Byte
        Dim CharacterRejectOff As Byte
        ' 16 bytes
        Dim ResultsReference As Byte
        Dim LexMode As Byte
        Dim OCRBOnly As Byte
        Dim Speed As Byte
        ' 20 bytes
        Dim FontStyleInfoOff As Byte
        Dim Reserved1 As Byte
        Dim Reserved2 As Byte
        Dim Reserved3 As Byte
        ' 24 bytes
        Dim CCAlgorithm As Int32
        ' 28 bytes
        Dim CCThreshold As Single
        ' 32 bytes
        Dim CGAlgorithm As Int32 'V5 addition
        ' 36 bytes
        Dim ExtraInfFlags As Int32
        ' 40 bytes
        <VBFixedArray(45), MarshalAs(UnmanagedType.ByValArray, SizeConst:=46)>
        Dim DisableLangs() As Byte
        Dim Reserved4 As Byte
        Dim Reserved5 As Byte
        ' 88 bytes
        'Line below: V5 extended from 601 to 608
        <VBFixedArray(607), MarshalAs(UnmanagedType.ByValArray, SizeConst:=608)>
        Public DisableCharW() As Byte
        ' 696 bytes

        Public Sub Initialize()
            ReDim DisableLangs(45)
            ReDim DisableCharW(607) 'V5 extended 600 to 607
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
        Dim DominantLanguage As Integer 'V5 addition
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSITEM_EG
        Dim Confidence As Single
        Dim StructId As UShort
        Dim OCRCharWUnicode As UShort 'V5 split from OCRChaW
        Dim OCRCharWInternal As UShort 'V5 split from OCRChaW
        Dim FontID As UShort
        Dim FontStyleInfo As UShort
        Dim XPos As UShort
        Dim YPos As UShort
        Dim XDim As UShort
        Dim YDim As UShort
        Dim YDimRef As UShort
        Dim Noise As UShort 'V5 addition
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTS_EG
        Dim Hdr As TOCRRESULTSHEADER_EG
        Dim Item() As TOCRRESULTSITEM_EG
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSITEMEXALT_EG
        Dim Factor As Single
        Dim Valid As UShort
        Dim OCRCharWUnicode As UShort 'V5 split from OCRChaW
        Dim OCRCharWInternal As UShort 'V5 split from OCRChaW
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure TOCRRESULTSITEMEX_EG
        Dim Confidence As Single
        Dim StructId As UShort
        Dim OCRCharWUnicode As UShort 'V5 split from OCRChaW
        Dim OCRCharWInternal As UShort 'V5 split from OCRChaW
        Dim FontID As UShort
        Dim FontStyleInfo As UShort
        Dim XPos As UShort
        Dim YPos As UShort
        Dim XDim As UShort
        Dim YDim As UShort
        Dim YDimRef As UShort
        Dim Noise As UShort 'V5 addition
        <VBFixedArray(4)> Dim Alt() As TOCRRESULTSITEMEXALT_EG  'N.B. this design reports the wrong structure size
        ' so we have to use careful marshaling when talking to the dll

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
#If SUPERSEDED Then
    <StructLayout(LayoutKind.Sequential)> _
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
        <VBFixedArray(255), MarshalAs(UnmanagedType.ByValArray, SizeConst:=256)> _
        Public DisableCharacter() As Short

        Public Sub Initialize()
            ReDim DisableCharacter(255)
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> _
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
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> _
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

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure TOCRRESULTSHEADER
        Dim StructId As Integer
        Dim XPixelsPerInch As Integer
        Dim YPixelsPerInch As Integer
        Dim NumItems As Integer
        Dim MeanConfidence As Single
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure TOCRRESULTSITEM
        Dim StructId As Short
        Dim OCRCha As Short
        Dim Confidence As Single
        Dim XPos As Short
        Dim YPos As Short
        Dim XDim As Short
        Dim YDim As Short
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure TOCRRESULTS
        Dim Hdr As TOCRRESULTSHEADER
        Dim Item() As TOCRRESULTSITEM
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure TOCRRESULTSITEMEXALT
        Dim Valid As Short
        Dim OCRCha As Short
        Dim Factor As Single
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
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

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure TOCRRESULTSEX
        Dim Hdr As TOCRRESULTSHEADER
        Dim Item() As TOCRRESULTSITEMEX
    End Structure
#End If
#End Region

#Region " Declares "
    ' 3 WAYS TO DECLARE THE FUNCTION
    'Declare Function testfn Lib "TOCRDLL" _
    '(<MarshalAs(UnmanagedType.LPWStr)> ByVal UniStr As String, ByVal ANsiStr As String, ByVal lens2 As Integer, ByRef ju As TOCRJOBINFO_EG) As Integer
    'Declare Ansi Function testfn Lib "TOCRDLL" _
    '(<MarshalAs(UnmanagedType.LPWStr)> ByVal UniStr As String, ByVal ANsiStr As String, ByVal lens2 As Integer, ByRef ju As TOCRJOBINFO_EG) As Integer
    'Declare Unicode Function testfn Lib "TOCRDLL" _
    '(ByVal UniStr As String, <MarshalAs(UnmanagedType.LPStr)> ByVal ANsiStr As String, ByVal lens2 As Integer, ByRef ju As TOCRJOBINFO_EG) As Integer

    'Release Win32 and Win64 version
    Declare Function TOCRInitialise Lib "TOCRDll" _
    (ByRef JobNo As Integer) As Integer

    Declare Function TOCRShutdown Lib "TOCRDll" _
        (ByVal JobNo As Integer) As Integer

    Declare Function TOCRDoJob_EG Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobInfo_EG As TOCRJOBINFO_EG) As Integer

    Declare Function TOCRDoJobPDF_EG Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobInfo_EG As TOCRJOBINFO_EG, ByVal Filename As String, ByRef PDFOpts As TOCRPROCESSPDFOPTIONS_EG) As Integer

    Declare Function TOCRWaitForJob Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobStatus As Integer) As Integer

    Declare Function TOCRWaitForAnyJob Lib "TOCRDll" _
        (ByRef WaitAnyStatus As Integer, ByRef JobNo As Integer) As Integer

    Declare Function TOCRGetJobDBInfo Lib "TOCRDll" _
        (ByVal JobSlotInf As System.IntPtr) As Integer

    Declare Function TOCRGetJobStatus Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobStatus As Integer) As Integer

    Declare Function TOCRGetJobStatusEx2 Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobStatus As Integer, ByRef Progress As Single, ByRef AutoOrientation As Integer, ByRef ExtraInfFlags As Integer) As Integer

    Declare Ansi Function TOCRGetJobStatusMsg Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Msg As System.Text.StringBuilder) As Integer

    Declare Ansi Function TOCRGetNumPages Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Filename As String, ByVal JobType As Integer, ByRef NumPages As Integer) As Integer

    Declare Function TOCRGetJobResultsEx_EG Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Mode As Integer, ByRef ResultsInf As Integer, ByVal Bytes As System.IntPtr) As Integer

    Declare Ansi Function TOCRGetLicenceInfoEx Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Licence As String, ByRef Volume As Integer, ByRef Time As Integer, ByRef Remaining As Integer, ByRef Features As Integer) As Integer

    Declare Ansi Function TOCRPopulateCharStatusMap Lib "TOCRDll" _
            (ByRef p_lang_opts As TOCRLANGUAGEOPTIONS_EG, ByRef p_usercharvalid As TOCRCHAROPTIONS_EG) As Integer

    ' These functions cannot be used to get/set the log file name in x64
    Declare Function TOCRSetConfig Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Parameter As Integer, ByVal Value As Integer) As Integer
    Declare Function TOCRGetConfig Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Parameter As Integer, ByRef Value As Integer) As Integer

    ' Convert a TIF or PDF file to a memory mapped file handle
    Declare Ansi Function TOCRConvertFormat Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal InputAddr As String, ByVal InputFormat As Integer, ByRef OutputAddr As System.IntPtr, ByVal OutputFormat As Integer, ByVal PageNo As Integer) As Integer

    ' Convert a TIF or PDF file to a bitmap file
    Declare Ansi Function TOCRConvertFormatHelper Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal InputAddr As String, ByVal InputFormat As Integer, ByVal OutputAddr As String, ByVal OutputFormat As Integer, ByVal PageNo As Integer) As Integer


    ' These functions can be used to get/set the log file name in x64
    Declare Ansi Function TOCRSetConfigStr Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Parameter As Integer, ByVal Value As String) As Integer
    Declare Ansi Function TOCRGetConfigStr Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Parameter As Integer, ByVal Value As System.Text.StringBuilder) As Integer
    ' Deprecated - use StringBuilder rather thsn String
    ' Declare Ansi Function TOCRGetConfigStr Lib "TOCRDll" _
    '    (ByVal JobNo As Integer, ByVal Parameter As Integer, ByVal Value As String) As Integer

    Declare Ansi Function TOCRGetFontName Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal FontID As Integer, ByVal FontName As System.Text.StringBuilder) As Integer

    Declare Ansi Function TOCRExtraInfGetMMF Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal ExtraInfFlag As Integer, ByRef MMF As System.IntPtr) As Integer

    Declare Function TOCRTWAINAcquire Lib "TOCRDll" _
        (ByRef NumberOfImages As Integer) As Integer

    Declare Function TOCRTWAINGetImages Lib "TOCRDll" _
        (ByVal GlobalMemoryDIBs As System.IntPtr) As Integer

    Declare Function TOCRTWAINSelectDS Lib "TOCRDll" _
        () As Integer

    Declare Function TOCRTWAINShowUI Lib "TOCRDll" _
        (ByVal Show As Short) As Integer

    ' Deprecated - use StringBuilder rather than String
    ' Declare Ansi Function TOCRGetJobStatusMsg Lib "TOCRDll" _
    '    (ByVal JobNo As Integer, ByVal Msg As String) As Integer

    ' Convert a TIF file to a bitmap file
    Declare Ansi Function TOCRConvertFormat Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal InputAddr As String, ByVal InputFormat As Integer, ByVal OutputAddr As String, ByVal OutputFormat As Integer, ByVal PageNo As Integer) As Integer


#End Region

#Region " SUPERSEDED Declares "
#If SUPERSEDED Then
    ' Superseded by TOCRGetConfig
    Declare Function TOCRGetErrorMode Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef ErrorMode As Integer) As Integer

    ' Superseded by TOCRSetConfig
    Declare Function TOCRSetErrorMode Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal ErrorMode As Integer) As Integer

    ' Superseded by TOCRDoJob_EG
    Declare Function TOCRDoJob2 Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobInfo As TOCRJOBINFO2) As Integer

    ' Superseded by TOCRDoJob2
    Declare Function TOCRDoJob Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobInfo As TOCRJOBINFO) As Integer

    ' Superseded by TOCRGetJobStatusEx2
    Declare Function TOCRGetJobStatusEx Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef JobStatus As Integer, ByRef Progress As Single, ByRef AutoOrientation As Integer) As Integer

    ' Superseded by TOCRGetJobResultsEx_EG
    Declare Function TOCRGetJobResults Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByRef ResultsInf As Integer, ByVal Bytes As System.IntPtr) As Integer

    ' Superseded by TOCRGetJobResultsEx_EG
    Declare Function TOCRGetJobResultsEx Lib "TOCRDll" _
        (ByVal JobNo As Integer, ByVal Mode As Integer, ByRef ResultsInf As Integer, ByVal Bytes As System.IntPtr) As Integer

    ' UNTESTED REDUNDANT - use the Bitmap class in .NET
    'Declare Function TOCRRotateMonoBitmap Lib "TOCRDll" _
    '    (ByRef hBmp As IntPtr, ByVal Width As Integer, ByVal Height As Integer, ByVal Orientation As Integer) As Integer

    ' UNTESTED - obsolete, use TOCRConvertFormat
    'Declare Ansi Function TOCRConvertTIFFtoDIB Lib "TOCRDll" _
    '    (ByVal JobNo As Integer, ByVal InputFilename As String, ByVal OutputFilename As String, ByVal PageNo As Integer) As Integer

    ' UNTESTED - Superseded by TOCRGetLicenceInfoEx
    'Declare Function TOCRGetLicenceInfo Lib "TOCRDll" _
    '    (ByRef NumOfJobSlots As Integer, ByRef Volume As Integer, ByRef Time As Integer, ByRef Remaining As Integer) As Integer
#End If
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
    Public Const TOCRJOBSTATUS_ERROR As Integer = -1    ' an error ocurred processing the last job
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
    Public Const TOCRJOBCC_HISTOGRAM As Integer = 9         ' Need equation
    Public Const TOCRJOBCC_REGIONS As Integer = 10          ' Need equation
    Public Const TOCRJOBCC_MEAN As Integer = 11             ' Need equation

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
    'Public Const TOCRRESULTSFONT_BOLD As UShort = 2     ' character is Bold - removed for v5
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
    'Public Const TOCRERR_FAILGETJOBSTATUS4 As Integer = 26

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

    Public Const TOCERR_TWAINPARTIALACQUIRE As Integer = 80
    'Public Const TOCERR_TWAINFAILEDACQUIRE As Integer = 81
    'Public Const TOCERR_TWAINNOIMAGES As Integer = 82
    'Public Const TOCERR_TWAINSELECTDSFAILED As Integer = 83
    'Public Const TOCERR_MMFNOTALLOWED As Integer = 84
    'Public Const TOCRERR_ILLEGALFONTID As Integer = 85

    'Public Const TOCRERR_FAILGETMMF As Integer = 90
    'Public Const TOCRERR_MMFNOTAVAILABLE As Integer = 91

    Public Const TOCRERR_PDFEXTRACTOR As Integer = 95
    Public Const TOCRERR_PDFERROR2 As Integer = 96
    Public Const TOCRERR_PDFARCHIVER As Integer = 97


    'Public Const TOCRERR_FONTSNOTLOADED As Integer = -2
#End Region

End Module
