' Transym OCR Demonstration program
'
' THE SOFTWARE IS PROVIDED "AS-IS" AND WITHOUT WARRANTY OF ANY KIND, 
' EXPRESS, IMPLIED OR OTHERWISE, INCLUDING WITHOUT LIMITATION, ANY 
' WARRANTY OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.  
'
' This program demonstrates calling TOCR version 4.0 from VB.NET.
'
' Copyright (C) 2012 Transym Computer Services Ltd.
'
'
' TOCR4.0DemoVB.NET Issue1

Option Strict On
Option Explicit On 

Imports System.Runtime.InteropServices
Imports System.Drawing.Imaging

Module Main
#If PLATFORM = "x64" Then
    Private mSample_TIF_file As String = Application.StartupPath & "\..\..\..\Sample.tif"
    Private mSample_BMP_file As String = Application.StartupPath & "\..\..\..\Sample.bmp"
#Else
    Private mSample_TIF_file As String = Application.StartupPath & "\..\..\Sample.tif"
    Private mSample_BMP_file As String = Application.StartupPath & "\..\..\Sample.bmp"
#End If

#Region " SDK Declares "
    Private Const DIB_RGB_COLORS As Integer = 0
    Private Const BI_RGB As Integer = 0
    Private Const BI_BITFIELDS As Integer = 3
    Private Const PAGE_READWRITE As Integer = 4
    Private Const FILE_MAP_WRITE As Integer = 2
    Private Const SRCCOPY As Integer = &HCC0020&

    <StructLayout(LayoutKind.Sequential, pack:=4)> _
    Structure RGBQUAD
        Dim rgbBlue As Byte
        Dim rgbGreen As Byte
        Dim rgbRed As Byte
        Dim rgbReserved As Byte
    End Structure ' RGBQUAD

    <StructLayout(LayoutKind.Sequential, pack:=4)> _
    Structure BITMAPINFOHEADER
        Dim biSize As Integer
        Dim biWidth As Integer
        Dim biHeight As Integer
        Dim biPlanes As Short
        Dim biBitCount As Short
        Dim biCompression As Integer
        Dim biSizeImage As Integer
        Dim biXPelsPerMeter As Integer
        Dim biYPelsPerMeter As Integer
        Dim biClrUsed As Integer
        Dim biClrImportant As Integer
    End Structure ' BITMAPINFOHEADER


    <StructLayout(LayoutKind.Sequential, pack:=4)> _
    Structure BITMAPINFO
        Dim bmih As BITMAPINFOHEADER
        <VBFixedArray(2), MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)> _
        Public cols As UInt32()
    End Structure ' BITMAPINFO

    Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Boolean
    Private Declare Function CreateFileMappingMy Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Integer, ByVal lpFileMappigAttributes As Integer, ByVal flProtect As Integer, ByVal dwMaximumSizeHigh As Integer, ByVal dwMaximumSizeLow As Integer, ByVal lpName As Integer) As IntPtr
    Private Declare Function MapViewOfFileMy Lib "kernel32" Alias "MapViewOfFile" (ByVal hFileMappingObject As IntPtr, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Integer, ByVal dwFileOffsetLow As Integer, ByVal dwNumberOfBytesToMap As Integer) As IntPtr
    Private Declare Function UnmapViewOfFileMy Lib "kernel32" Alias "UnmapViewOfFile" (ByVal lpBaseAddress As IntPtr) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As Integer, ByVal lpvSrc As IntPtr, ByVal cbCopy As Integer)
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As IntPtr) As IntPtr
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As IntPtr) As Integer
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As IntPtr) As Integer

    Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As IntPtr) As Boolean
    Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As IntPtr) As IntPtr
    Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hRefDC As IntPtr) As IntPtr

    Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Boolean
    Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As IntPtr) As Boolean
    Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As IntPtr, ByVal hObject As IntPtr) As IntPtr
    Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hdc As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As Integer) As Boolean
    Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As IntPtr, ByRef pbmi As BITMAPINFO, ByVal iUsage As UInt32, ByRef ppvBits As IntPtr, ByVal hSection As IntPtr, ByVal dwOffset As UInt32) As IntPtr

#End Region


    Sub Main()

        Example1() ' Demonstrates how to OCR a file
        Example2() ' Demonstrates how to OCR multiple files
        Example3() ' Demonstrates how to OCR an image using a memory mapped file created by TOCR
        Example4() ' Demonstrates how to OCR an image using a memory mapped file created here
        Example5() ' Retrieve information on Job Slot usage
        Example6() ' Retrieve information on Job Slots
        Example7() ' Get images from a TWAIN compatible device
        Example8() ' Demonstrates TOCRSetConfig and TOCRGetConfig

    End Sub

    ' Demonstrates how to OCR a file
    Private Sub Example1()

        Dim Status As Integer
        Dim JobNo As Integer
        Dim JobInfo2 As New TOCRJOBINFO2
        Dim Answer As String = ""
        Dim Results As New TOCRRESULTS

        TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX)

        JobInfo2.Initialize()

        JobInfo2.InputFile = mSample_TIF_file
        JobInfo2.JobType = TOCRJOBTYPE_TIFFFILE

        ' or
        'JobInfo2.InputFile = mSample_BMP_file
        'JobInfo2.JobType = TOCRJOBTYPE_DIBFILE

        Status = TOCRInitialise(JobNo)
        If Status = TOCR_OK Then
            ' or
            'If OCRPoll(JobNo, JobInfo2) Then
            If OCRWait(JobNo, JobInfo2) Then
                If GetResults(JobNo, Results) Then
                    If FormatResults(Results, Answer) Then
                        MsgBox(Answer, MsgBoxStyle.Information, "Example 1")
                    End If
                End If
            End If
            TOCRShutdown(JobNo)
        End If

    End Sub ' Example 1

    ' Demonstrates how to OCR multiple files
    Private Sub Example2()

        Dim Status As Integer
        Dim JobNo As Integer
        Dim JobInfo2 As New TOCRJOBINFO2
        Dim Results As New TOCRRESULTS
        Dim CountDone As Integer = 0

        TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX)

        JobInfo2.Initialize()

        Status = TOCRInitialise(JobNo)
        If Status = TOCR_OK Then

            ' 1st file
            JobInfo2.InputFile = mSample_TIF_file
            JobInfo2.JobType = TOCRJOBTYPE_TIFFFILE
            If OCRWait(JobNo, JobInfo2) Then
                If GetResults(JobNo, Results) Then
                    CountDone += 1
                End If
            End If

            ' 2nd file
            JobInfo2.InputFile = mSample_BMP_file
            JobInfo2.JobType = TOCRJOBTYPE_DIBFILE
            If OCRWait(JobNo, JobInfo2) Then
                If GetResults(JobNo, Results) Then
                    CountDone += 1
                End If
            End If
            TOCRShutdown(JobNo)
        End If

        MsgBox(CountDone.ToString() & " of 2 jobs done", MsgBoxStyle.Information, "Example 2")
    End Sub ' Example 2

    ' Demonstrates how to OCR an image using a memory mapped file created by TOCR
    Private Sub Example3()

        Dim Status As Integer
        Dim JobNo As Integer
        Dim Answer As String = ""
        Dim MMFhandle As IntPtr = IntPtr.Zero
        Dim Results As New TOCRRESULTSEX
        Dim JobInfo2 As New TOCRJOBINFO2

        TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX)

        JobInfo2.Initialize()

        JobInfo2.JobType = TOCRJOBTYPE_MMFILEHANDLE

        Status = TOCRInitialise(JobNo)
        If Status = TOCR_OK Then
            Status = TOCRConvertFormat(JobNo, mSample_TIF_file, TOCRCONVERTFORMAT_TIFFFILE, MMFhandle, TOCRCONVERTFORMAT_MMFILEHANDLE, 0)
            If Status = TOCR_OK Then
                JobInfo2.hMMF = MMFhandle
                If OCRWait(JobNo, JobInfo2) Then
                    If GetResults(JobNo, Results) Then
                        If FormatResults(Results, Answer) Then
                            MsgBox(Answer, MsgBoxStyle.Information, "Example 3")
                        End If
                    End If
                End If
            End If
            If Not MMFhandle.Equals(IntPtr.Zero) Then
                CloseHandle(MMFhandle)
            End If
            TOCRShutdown(JobNo)
        End If

    End Sub ' Example 3

    ' Demonstrates how to OCR an image using a memory mapped file created here
    Private Sub Example4()

        Dim BMP As Bitmap
        Dim Status As Integer
        Dim JobNo As Integer
        Dim JobInfo2 As New TOCRJOBINFO2
        Dim Answer As String = ""
        Dim MMFhandle As IntPtr = IntPtr.Zero
        Dim Results As New TOCRRESULTS

        BMP = New Bitmap(mSample_BMP_file)

        MMFhandle = ConvertBitmapToMMF(BMP)
        'MMFhandle = ConvertBitmapToMMF2(BMP)

        If Not MMFhandle.Equals(IntPtr.Zero) Then

            TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX)

            JobInfo2.Initialize()

            JobInfo2.JobType = TOCRJOBTYPE_MMFILEHANDLE

            Status = TOCRInitialise(JobNo)

            If Status = TOCR_OK Then
                JobInfo2.hMMF = MMFhandle

                If OCRWait(JobNo, JobInfo2) Then
                    If GetResults(JobNo, Results) Then
                        If FormatResults(Results, Answer) Then
                            MsgBox(Answer, MsgBoxStyle.Information, "Example 4")
                        End If
                    End If
                End If
                TOCRShutdown(JobNo)
            End If
            CloseHandle(MMFhandle)
        End If

    End Sub ' Example 4

    ' Retrieve information on Job Slot usage
    Private Sub Example5()

        Dim Status As Integer
        Dim NumSlots As Integer
        Dim JobSlotInf() As Integer
        Dim SlotNo As Integer
        Dim Msg As String
        Dim BytesGC As GCHandle

        ' uncomment to see effect on usage
        'Dim JobNo As Integer
        'Status = TOCRInitialise(JobNo)

        NumSlots = TOCRGetJobDBInfo(IntPtr.Zero)
        If NumSlots > 0 Then
            ReDim JobSlotInf(NumSlots - 1)
            BytesGC = GCHandle.Alloc(JobSlotInf, GCHandleType.Pinned)
            Status = TOCRGetJobDBInfo(BytesGC.AddrOfPinnedObject)
            BytesGC.Free()
            If Status = TOCR_OK Then
                Msg = "Slot usage is" & vbCrLf
                For SlotNo = 0 To NumSlots - 1
                    Msg = Msg & vbCrLf & "Slot" & Str$(SlotNo) & " is "
                    Select Case JobSlotInf(SlotNo)
                        Case TOCRJOBSLOT_FREE
                            Msg = Msg & "free"
                        Case TOCRJOBSLOT_OWNEDBYYOU
                            Msg = Msg & "owned by you"
                        Case TOCRJOBSLOT_BLOCKEDBYYOU
                            Msg = Msg & "blocked by you"
                        Case TOCRJOBSLOT_OWNEDBYOTHER
                            Msg = Msg & "owned by another process"
                        Case TOCRJOBSLOT_BLOCKEDBYOTHER
                            Msg = Msg & "blocked by another process"
                    End Select
                Next SlotNo
                MsgBox(Msg, MsgBoxStyle.Information, "Example 5")
            End If ' Status = TOCR_OK
        Else
            MsgBox("No Job Slots", MsgBoxStyle.Critical, "Example 5")
        End If ' NumSlots > 0

        'TOCRShutdown(JobNo)

    End Sub ' Example 5

    ' Retrieve information on Job Slots
    Private Sub Example6()

        Dim Status As Integer
        Dim NumSlots As Integer
        Dim SlotNo As Integer
        Dim Msg As String
        Dim Volume As Integer
        Dim Time As Integer
        Dim Remaining As Integer
        Dim Features As Integer
        Dim Licence As String = Space$(20)

        NumSlots = TOCRGetJobDBInfo(IntPtr.Zero)
        If NumSlots > 0 Then
            Msg = "Slot usage is" & vbCrLf
            For SlotNo = 0 To NumSlots - 1
                Msg = Msg & vbCrLf & "Slot" & Str$(SlotNo)
                Status = TOCRGetLicenceInfoEx(SlotNo, Licence, Volume, Time, Remaining, Features)
                If Status = TOCR_OK Then
                    Msg = Msg & " " & Left$(Licence, 19)
                    Select Case Features
                        Case TOCRLICENCE_STANDARD
                            Msg = Msg & " STANDARD licence"
                        Case TOCRLICENCE_EURO
                            If Licence.ToString() = "5AD4-1D96-F632-8912" Then
                                Msg = Msg & " EURO TRIAL licence"
                            Else
                                Msg = Msg & " EURO licence"
                            End If
                        Case TOCRLICENCE_EUROUPGRADE
                            Msg = Msg & " EURO UPGRADE licence"
                        Case TOCRLICENCE_V3SE
                            If Licence.ToString() = "2E72-2B35-643A-0851" Then
                                Msg = Msg & " V3 TRIAL licence"
                            Else
                                Msg = Msg & " V3 licence"
                            End If
                        Case TOCRLICENCE_V3SEUPGRADE
                            Msg = Msg & " V1/2 UPGRADE to V3 SE licence"
                        Case TOCRLICENCE_V3PRO
                            Msg = Msg & " V3 Pro/V4 licence"
                        Case TOCRLICENCE_V3PROUPGRADE
                            Msg = Msg & " V1/2 UPGRADE to V3 Pro/V4 licence"
                        Case TOCRLICENCE_V3SEPROUPGRADE
                            Msg = Msg & " V3 SE UPGRADE to V3 Pro/V4 licence"
                    End Select
                    If Volume <> 0 Or Time <> 0 Then
                        Msg = Msg & Str$(Remaining)
                        If Time <> 0 Then
                            Msg = Msg & " days"
                        Else
                            Msg = Msg & " A4 pages"
                        End If
                        Msg = Msg & " remaining on licence"
                    End If
                End If
            Next SlotNo
            MsgBox(Msg, MsgBoxStyle.Information, "Example 6")
        Else
            MsgBox("No Job Slots", MsgBoxStyle.Critical, "Example 6")
        End If ' NumSlots > 0

    End Sub ' Example 6

    ' Get images from a TWAIN compatible device
    Private Sub Example7()

        Dim Status As Integer
        Dim NumImages As Integer
        Dim CntImages As Integer
        Dim hDIBs() As IntPtr
        Dim BytesGC As GCHandle
        Dim BMP As Bitmap
        Dim ImgNo As Integer

        Status = TOCRTWAINSelectDS() ' select the TWAIN device
        If Status = TOCR_OK Then
            Status = TOCRTWAINShowUI(1)
            Status = TOCRTWAINAcquire(NumImages)
            If Status = TOCR_OK And NumImages > 0 Then

                ReDim hDIBs(NumImages - 1)
                BytesGC = GCHandle.Alloc(hDIBs, GCHandleType.Pinned)
                Status = TOCRTWAINGetImages(BytesGC.AddrOfPinnedObject())
                BytesGC.Free()

                For ImgNo = 0 To NumImages - 1

                    ' Convert the memory block to a bitmap.  If you do not require a bitmap
                    ' you could convert the memory block to a memory mapped file and OCR it.
                    BMP = ConvertMemoryBlockToBitmap(hDIBs(ImgNo))

                    ' You could combine this code with Example2 to OCR the bitmap
                    If Not BMP.Equals(Nothing) Then
                        ' could save the bitmap here
                        'BMP.Save("A.bmp", System.Drawing.Imaging.ImageFormat.Bmp)
                        BMP.Dispose()
                        BMP = Nothing
                        CntImages += 1
                    End If

                    ' Free the global memory block
                    Status = GlobalFree(hDIBs(ImgNo))
                Next ImgNo
            End If
        End If
        MsgBox(CntImages.ToString() & " images successfully received", MsgBoxStyle.Information, "Example 7")

    End Sub ' Example 7

    ' Demonstrates TOCRSetConfig and TOCRGetConfig
    Private Sub Example8()

        Dim JobNo As Integer
        Dim Msg As String
        Dim Value As Integer

        Msg = Space(250)

        ' Override the INI file settings for all new jobs
        TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX)
        TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_SRV_ERRORMODE, TOCRERRORMODE_MSGBOX)

        TOCRGetConfigStr(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_LOGFILE, Msg)
        MsgBox("Default Log file name = " & Msg, MsgBoxStyle.Information, "Example 8")

        TOCRSetConfigStr(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_LOGFILE, "Loggederrs.lis")
        TOCRGetConfigStr(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_LOGFILE, Msg)
        MsgBox("New default Log file name = " & Msg, MsgBoxStyle.Information, "Example 8")

        TOCRInitialise(JobNo)
        TOCRSetConfig(JobNo, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_NONE)

        TOCRGetConfig(JobNo, TOCRCONFIG_DLL_ERRORMODE, Value)
        MsgBox("Job DLL error mode = " & Value.ToString(), MsgBoxStyle.Information, "Example 8")

        TOCRGetConfig(JobNo, TOCRCONFIG_SRV_ERRORMODE, Value)
        MsgBox("Job Service error mode = " & Value.ToString(), MsgBoxStyle.Information, "Example 8")

        ' Cause an error to be sent to Loggederrs.lis
        TOCRSetConfig(JobNo, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_LOG)
        TOCRSetConfig(JobNo, 1000, TOCRERRORMODE_LOG)

        TOCRShutdown(JobNo)

    End Sub ' Example 8

    ' Wait for the engine to complete
    Private Function OCRWait(ByVal JobNo As Integer, ByVal JobInfo2 As TOCRJOBINFO2) As Boolean

        Dim Status As Integer
        Dim JobStatus As Integer
        Dim Msg As String
        Dim ErrorMode As Integer


        Status = TOCRDoJob2(JobNo, JobInfo2)
        If Status = TOCR_OK Then
            Status = TOCRWaitForJob(JobNo, JobStatus)
        End If

        If Status = TOCR_OK And JobStatus = TOCRJOBSTATUS_DONE Then
            OCRWait = True
        Else
            OCRWait = False

            ' If something hass gone wrong display a message
            ' (Check that the OCR engine hasn't already displayed a message)
            TOCRGetConfig(JobNo, TOCRCONFIG_DLL_ERRORMODE, ErrorMode)
            If ErrorMode = TOCRERRORMODE_NONE Then
                Msg = Space$(TOCRJOBMSGLENGTH)
                TOCRGetJobStatusMsg(JobNo, Msg)
                MsgBox(Msg, MsgBoxStyle.Critical, "OCRWait")
            End If
        End If

    End Function

    ' Wait for the engine to complete by polling
    Private Function OCRPoll(ByVal JobNo As Integer, ByVal JobInfo2 As TOCRJOBINFO2) As Boolean

        Dim Status As Integer
        Dim JobStatus As Integer
        Dim Msg As String
        Dim ErrorMode As Integer
        Dim Progress As Single
        Dim AutoOrientation As Integer

        Status = TOCRDoJob2(JobNo, JobInfo2)
        If Status = TOCR_OK Then
            Do
                'Status = TOCRGetJobStatus(JobNo, JobStatus)
                Status = TOCRGetJobStatusEx(JobNo, JobStatus, Progress, AutoOrientation)

                ' Do something whilst the OCR engine runs
                Application.DoEvents() : Threading.Thread.Sleep(100) : Application.DoEvents()
                Console.WriteLine("Progress" & Str$(Int(Progress * 100)) & "%")

            Loop While Status = TOCR_OK And JobStatus = TOCRJOBSTATUS_BUSY
        End If

        If Status = TOCR_OK And JobStatus = TOCRJOBSTATUS_DONE Then
            OCRPoll = True
        Else
            OCRPoll = False

            ' If something hass gone wrong display a message
            ' (Check that the OCR engine hasn't already displayed a message)
            TOCRGetConfig(JobNo, TOCRCONFIG_DLL_ERRORMODE, ErrorMode)
            If ErrorMode = TOCRERRORMODE_NONE Then
                Msg = Space$(TOCRJOBMSGLENGTH)
                TOCRGetJobStatusMsg(JobNo, Msg)
                MsgBox(Msg, MsgBoxStyle.Critical, "OCRPoll")
            End If
        End If

    End Function

    ' OVERLOADED function to retrieve the results from the service process and load into 'Results'
    ' Remember the character numbers returned refer to the Windows character set.
    Private Function GetResults(ByVal JobNo As Integer, ByRef Results As TOCRRESULTS) As Boolean

        Dim ResultsInf As Integer ' number of bytes needed for results
        Dim BytesGC As GCHandle
        Dim AddrOfItemBytes As System.IntPtr
        Dim ItemNo As Integer
        Dim Bytes() As Byte
        Dim Offset As Integer

        GetResults = False
        Results.Hdr.NumItems = 0

        If TOCRGetJobResults(JobNo, ResultsInf, IntPtr.Zero) = TOCR_OK Then
            If ResultsInf > 0 Then
                ReDim Bytes(ResultsInf - 1)
                Dim s As Integer
                For s = 0 To ResultsInf - 1 : Bytes(s) = 255 : Next s

                ' pin the Bytes array so that TOCRGetJobResults can write to it
                BytesGC = GCHandle.Alloc(Bytes, GCHandleType.Pinned)

                If TOCRGetJobResults(JobNo, ResultsInf, BytesGC.AddrOfPinnedObject) = TOCR_OK Then
                    With Results
                        .Hdr = CType(Marshal.PtrToStructure(BytesGC.AddrOfPinnedObject, GetType(TOCRRESULTSHEADER)), TOCRRESULTSHEADER)
                        If .Hdr.NumItems > 0 Then
                            ReDim .Item(.Hdr.NumItems - 1)
                            Offset = Marshal.SizeOf(GetType(TOCRRESULTSHEADER))
                            For ItemNo = 0 To .Hdr.NumItems - 1
                                AddrOfItemBytes = Marshal.UnsafeAddrOfPinnedArrayElement(Bytes, Offset)
                                .Item(ItemNo) = CType(Marshal.PtrToStructure(AddrOfItemBytes, GetType(TOCRRESULTSITEM)), TOCRRESULTSITEM)
                                Offset = Offset + Marshal.SizeOf(GetType(TOCRRESULTSITEM))
                            Next ItemNo
                        End If ' .Hdr.NumItems > 0

                        GetResults = True

                    End With ' results
                End If ' TOCRGetJobResults(JobNo, ResultsInf, Bytes(0)) = TOCR_OK

                BytesGC.Free()

            End If ' ResultsInf > 0
        End If ' TOCRGetJobResults(JobNo, ResultsInf, 0) = TOCR_OK

    End Function

    ' copy of TOCRRESULTSITEMEX without the Alt[] array 
    Structure TOCRRESULTSITEMEXHDR
        Dim StructId As Short
        Dim OCRCha As Short
        Dim Confidence As Single
        Dim XPos As Short
        Dim YPos As Short
        Dim XDim As Short
        Dim YDim As Short
    End Structure

    ' OVERLOADED function to retrieve the results from the service process and load into 'ResultsEx'
    ' Remember the character numbers returned refer to the Windows character set.
    Private Function GetResults(ByVal JobNo As Integer, ByRef ResultsEx As TOCRRESULTSEX) As Boolean

        Dim ResultsInf As Integer ' number of bytes needed for results
        Dim BytesGC As GCHandle
        Dim AddrOfItemBytes As System.IntPtr
        Dim ItemNo As Integer
        Dim AltNo As Integer
        Dim Bytes() As Byte
        Dim Offset As Integer
        Dim ItemHdr As TOCRRESULTSITEMEXHDR


        GetResults = False
        ResultsEx.Hdr.NumItems = 0

        If TOCRGetJobResultsEx(JobNo, TOCRGETRESULTS_EXTENDED, ResultsInf, IntPtr.Zero) = TOCR_OK Then
            If ResultsInf > 0 Then
                ReDim Bytes(ResultsInf - 1)
                ' pin the Bytes array so that TOCRGetJobResultsEx can write to it
                BytesGC = GCHandle.Alloc(Bytes, GCHandleType.Pinned)

                If TOCRGetJobResultsEx(JobNo, TOCRGETRESULTS_EXTENDED, ResultsInf, BytesGC.AddrOfPinnedObject) = TOCR_OK Then
                    With ResultsEx
                        .Hdr = CType(Marshal.PtrToStructure(BytesGC.AddrOfPinnedObject, GetType(TOCRRESULTSHEADER)), TOCRRESULTSHEADER)
                        If .Hdr.NumItems > 0 Then
                            ReDim .Item(.Hdr.NumItems - 1)
                            Offset = Marshal.SizeOf(GetType(TOCRRESULTSHEADER))
                            For ItemNo = 0 To .Hdr.NumItems - 1
                                AddrOfItemBytes = Marshal.UnsafeAddrOfPinnedArrayElement(Bytes, Offset)

                                ' Cannot Marshal TOCRRESULTSITEMEX so use copy of structure header
                                ' This unfortunately means a double copy of the data
                                ItemHdr = CType(Marshal.PtrToStructure(AddrOfItemBytes, GetType(TOCRRESULTSITEMEXHDR)), TOCRRESULTSITEMEXHDR)
                                With .Item(ItemNo)
                                    .Initialize()
                                    .StructId = ItemHdr.StructId
                                    .OCRCha = ItemHdr.OCRCha
                                    .Confidence = ItemHdr.Confidence
                                    .XPos = ItemHdr.XPos
                                    .YPos = ItemHdr.YPos
                                    .XDim = ItemHdr.XDim
                                    .YDim = ItemHdr.YDim
                                    Offset = Offset + Marshal.SizeOf(GetType(TOCRRESULTSITEMEXHDR))
                                    For AltNo = 0 To 4
                                        AddrOfItemBytes = Marshal.UnsafeAddrOfPinnedArrayElement(Bytes, Offset)
                                        .Alt(AltNo) = CType(Marshal.PtrToStructure(AddrOfItemBytes, GetType(TOCRRESULTSITEMEXALT)), TOCRRESULTSITEMEXALT)
                                        Offset = Offset + Marshal.SizeOf(GetType(TOCRRESULTSITEMEXALT))
                                    Next AltNo
                                End With
                            Next ItemNo
                        End If ' .Hdr.NumItems > 0

                        GetResults = True

                    End With ' results
                End If ' TOCRGetJobResults(JobNo, ResultsInf, Bytes(0)) = TOCR_OK

                BytesGC.Free()

            End If ' ResultsInf > 0
        End If ' TOCRGetJobResults(JobNo, ResultsInf, 0) = TOCR_OK

    End Function

    'OVERLOADED function to convert results to a string
    Private Function FormatResults(ByVal Results As TOCRRESULTS, ByRef Answer As String) As Boolean

        Dim ItemNo As Integer

        FormatResults = False
        Answer = ""

        With Results
            If .Hdr.NumItems > 0 Then
                For ItemNo = 0 To .Hdr.NumItems - 1
                    If Chr(.Item(ItemNo).OCRCha) = vbCr Then
                        Answer = Answer & vbCrLf
                    Else
                        Answer = Answer & Chr(.Item(ItemNo).OCRCha)
                    End If
                Next ItemNo
                FormatResults = True
            Else
                MsgBox("No results returned", MsgBoxStyle.Information, "FormatResults")
            End If
        End With

    End Function

    'OVERLOADED function to convert results to a string
    Private Function FormatResults(ByVal ResultsEx As TOCRRESULTSEX, ByRef Answer As String) As Boolean

        Dim ItemNo As Integer

        FormatResults = False

        With ResultsEx
            If .Hdr.NumItems > 0 Then
                For ItemNo = 0 To .Hdr.NumItems - 1
                    If Chr(.Item(ItemNo).OCRCha) = vbCr Then
                        Answer = Answer & vbCrLf
                    Else
                        Answer = Answer & Chr(.Item(ItemNo).OCRCha)
                    End If
                Next ItemNo
                FormatResults = True
            Else
                MsgBox("No results returned", MsgBoxStyle.Information, "FormatResults")
            End If
        End With

    End Function

    ' Convert a bitmap to 1bpp
    Private Function ConvertTo1bpp(ByVal BMPIn As Bitmap) As Bitmap

        Dim bmi As New BITMAPINFO
        Dim hbmIn As IntPtr = BMPIn.GetHbitmap()

        bmi.bmih.biSize = CInt(Marshal.SizeOf(bmi.bmih))
        bmi.bmih.biWidth = BMPIn.Width
        bmi.bmih.biHeight = BMPIn.Height
        bmi.bmih.biPlanes = 1
        bmi.bmih.biBitCount = 1
        bmi.bmih.biCompression = BI_RGB
        bmi.bmih.biSizeImage = CInt((((BMPIn.Width + 7) And &HFFFFFFF8&) >> 3) * BMPIn.Height)
        bmi.bmih.biXPelsPerMeter = System.Convert.ToInt32(BMPIn.HorizontalResolution * 100 / 2.54)
        bmi.bmih.biYPelsPerMeter = System.Convert.ToInt32(BMPIn.VerticalResolution * 100 / 2.54)
        bmi.bmih.biClrUsed = 2
        bmi.bmih.biClrImportant = 2
        ReDim bmi.cols(1)  ' see the definition of BITMAPINFO()
        bmi.cols(0) = Convert.ToUInt32(0)
        bmi.cols(1) = Convert.ToUInt32(&HFFFFFF)

        Dim dummy As IntPtr
        Dim hbm As IntPtr = CreateDIBSection(IntPtr.Zero, bmi, Convert.ToUInt32(DIB_RGB_COLORS), dummy, IntPtr.Zero, Convert.ToUInt32(0))

        Dim scrnDC As IntPtr = GetDC(IntPtr.Zero)
        Dim hDCIn As IntPtr = CreateCompatibleDC(scrnDC)

        SelectObject(hDCIn, hbmIn)
        Dim hDC As IntPtr = CreateCompatibleDC(scrnDC)
        SelectObject(hDC, hbm)

        BitBlt(hDC, 0, 0, BMPIn.Width, BMPIn.Height, hDCIn, 0, 0, SRCCOPY)

        Dim BMP As Bitmap = Bitmap.FromHbitmap(hbm)

        DeleteDC(hDCIn)
        DeleteDC(hDC)
        ReleaseDC(IntPtr.Zero, scrnDC)
        DeleteObject(hbmIn)
        DeleteObject(hbm)

        Return BMP

    End Function

    ' Convert a bitmap to a memory mapped file.
    ' It does this by constructing a GDI bitmap in a byte array and copying this to a memory mapped file.
    Private Function ConvertBitmapToMMF(ByVal BMPIn As Bitmap, _
        Optional ByVal DiscardBitmap As Boolean = True, _
        Optional ByVal ConvertTo1Bit As Boolean = True) As IntPtr

        Dim BMP As Bitmap
        Dim BIH As BITMAPINFOHEADER
        Dim BMPData As BitmapData
        Dim ImageSize As Integer
        Dim Bytes() As Byte
        Dim BytesGC As GCHandle
        Dim MMFsize As Integer
        Dim PalEntries As Integer
        Dim PalEntry As Integer
        Dim rgb As RGBQUAD
        Dim Offset As Integer
        Dim MMFhandle As IntPtr = IntPtr.Zero
        Dim MMFview As IntPtr = IntPtr.Zero

        ConvertBitmapToMMF = IntPtr.Zero

        If DiscardBitmap Then   ' can destroy input bitmap
            If ConvertTo1Bit Then
                BMP = ConvertTo1bpp(BMPIn)
                BMPIn.Dispose()
                BMPIn = Nothing
            Else
                BMP = BMPIn
            End If
        Else                    ' must keep input bitmap unchanged
            If ConvertTo1Bit Then
                BMP = ConvertTo1bpp(BMPIn)
            Else
                BMP = BMPIn.Clone(New Rectangle(New Point, BMPIn.Size), BMPIn.PixelFormat)
            End If
        End If

        ' Flip the bitmap (GDI+ bitmap scan lines are top down, GDI are bottom up)
        BMP.RotateFlip(RotateFlipType.RotateNoneFlipY)

        BMPData = BMP.LockBits(New Rectangle(New Point, BMP.Size), ImageLockMode.ReadOnly, BMP.PixelFormat)
        ImageSize = BMPData.Stride * BMP.Height

        PalEntries = BMP.Palette.Entries.Length

        BIH.biWidth = BMP.Width
        BIH.biHeight = BMP.Height
        BIH.biPlanes = 1
        BIH.biSize = Marshal.SizeOf(BIH)
        BIH.biClrImportant = 0
        BIH.biCompression = BI_RGB
        BIH.biSizeImage = ImageSize
        BIH.biXPelsPerMeter = CInt(BMP.HorizontalResolution * 100 / 2.54)
        BIH.biYPelsPerMeter = CInt(BMP.VerticalResolution * 100 / 2.54)

        ' Most of these formats are untested and the alpha channel is ignored
        Select Case BMP.PixelFormat
            Case PixelFormat.Format1bppIndexed
                BIH.biBitCount = 1
            Case PixelFormat.Format4bppIndexed
                BIH.biBitCount = 4
            Case PixelFormat.Format8bppIndexed
                BIH.biBitCount = 8
            Case PixelFormat.Format16bppArgb1555, PixelFormat.Format16bppGrayScale, PixelFormat.Format16bppRgb555, PixelFormat.Format16bppRgb565
                BIH.biBitCount = 16
                PalEntries = 0
            Case PixelFormat.Format24bppRgb
                BIH.biBitCount = 24
                PalEntries = 0
            Case PixelFormat.Format32bppArgb, PixelFormat.Format32bppPArgb, PixelFormat.Format32bppRgb
                BIH.biBitCount = 32
                PalEntries = 0
        End Select
        BIH.biClrUsed = PalEntries

        MMFsize = Marshal.SizeOf(BIH) + PalEntries * Marshal.SizeOf(GetType(RGBQUAD)) + ImageSize
        ReDim Bytes(MMFsize)

        BytesGC = GCHandle.Alloc(Bytes, GCHandleType.Pinned)
        Marshal.StructureToPtr(BIH, BytesGC.AddrOfPinnedObject, True)
        Offset = Marshal.SizeOf(BIH)
        For PalEntry = 0 To PalEntries - 1
            rgb.rgbRed = BMP.Palette.Entries(PalEntry).R
            rgb.rgbGreen = BMP.Palette.Entries(PalEntry).G
            rgb.rgbBlue = BMP.Palette.Entries(PalEntry).B
            Marshal.StructureToPtr(rgb, Marshal.UnsafeAddrOfPinnedArrayElement(Bytes, Offset), False)
            Offset = Offset + Marshal.SizeOf(rgb)
        Next
        BytesGC.Free()
        Marshal.Copy(BMPData.Scan0, Bytes, Offset, ImageSize)
        BMP.UnlockBits(BMPData)
        BMPData = Nothing
        BMP.Dispose()
        BMP = Nothing

        MMFhandle = CreateFileMappingMy(&HFFFFFFFF, 0&, PAGE_READWRITE, 0, MMFsize, 0&)
        If Not MMFhandle.Equals(IntPtr.Zero) Then
            MMFview = MapViewOfFileMy(MMFhandle, FILE_MAP_WRITE, 0, 0, 0)
            If MMFview.Equals(IntPtr.Zero) Then
                CloseHandle(MMFhandle)
            Else
                Marshal.Copy(Bytes, 0, MMFview, MMFsize)
                UnmapViewOfFileMy(MMFview)
                ConvertBitmapToMMF = MMFhandle
            End If
        End If

        Bytes = Nothing

        If MMFhandle.Equals(IntPtr.Zero) Then
            MsgBox("Failed to convert bitmap", MsgBoxStyle.Critical, "ConvertBitmapToMMF")
        End If

    End Function

    ' Convert a bitmap to a memory mapped file
    ' (Same as ConvertBitmapToMMF but uses CopyMemory to avoid using a byte array)
    Private Function ConvertBitmapToMMF2(ByRef BMPIn As Bitmap, _
        Optional ByVal DiscardBitmap As Boolean = True, _
        Optional ByVal ConvertTo1Bit As Boolean = True) As IntPtr

        Dim BMP As Bitmap
        Dim BIH As BITMAPINFOHEADER
        Dim BMPData As BitmapData
        Dim ImageSize As Integer
        Dim MMFsize As Integer
        Dim PalEntries As Integer
        Dim PalEntry As Integer
        Dim rgb As RGBQUAD
        Dim rgbGC As GCHandle
        Dim Offset As Integer
        Dim MMFhandle As IntPtr = IntPtr.Zero
        Dim MMFview As IntPtr = IntPtr.Zero

        ConvertBitmapToMMF2 = IntPtr.Zero

        If DiscardBitmap Then   ' can destroy input bitmap
            If ConvertTo1Bit Then
                BMP = ConvertTo1bpp(BMPIn)
                BMPIn.Dispose()
                BMPIn = Nothing
            Else
                BMP = BMPIn
            End If
        Else                    ' must keep input bitmap unchanged
            If ConvertTo1Bit Then
                BMP = ConvertTo1bpp(BMPIn)
            Else
                BMP = BMPIn.Clone(New Rectangle(New Point, BMPIn.Size), BMPIn.PixelFormat)
            End If
        End If

        ' Flip the bitmap (GDI+ bitmap scan lines are top down, GDI are bottom up)
        BMP.RotateFlip(RotateFlipType.RotateNoneFlipY)

        BMPData = BMP.LockBits(New Rectangle(New Point, BMP.Size), ImageLockMode.ReadOnly, BMP.PixelFormat)
        ImageSize = BMPData.Stride * BMP.Height

        PalEntries = BMP.Palette.Entries.Length

        BIH.biWidth = BMP.Width
        BIH.biHeight = BMP.Height
        BIH.biPlanes = 1
        BIH.biSize = Marshal.SizeOf(BIH)
        BIH.biClrImportant = 0
        BIH.biCompression = BI_RGB
        BIH.biSizeImage = ImageSize
        BIH.biXPelsPerMeter = CInt(BMP.HorizontalResolution * 100 / 2.54)
        BIH.biYPelsPerMeter = CInt(BMP.VerticalResolution * 100 / 2.54)

        ' Most of these formats are untested and the alpha channel is ignored
        Select Case BMP.PixelFormat
            Case PixelFormat.Format1bppIndexed
                BIH.biBitCount = 1
            Case PixelFormat.Format4bppIndexed
                BIH.biBitCount = 4
            Case PixelFormat.Format8bppIndexed
                BIH.biBitCount = 8
            Case PixelFormat.Format16bppArgb1555, PixelFormat.Format16bppGrayScale, PixelFormat.Format16bppRgb555, PixelFormat.Format16bppRgb565
                BIH.biBitCount = 16
                PalEntries = 0
            Case PixelFormat.Format24bppRgb
                BIH.biBitCount = 24
                PalEntries = 0
            Case PixelFormat.Format32bppArgb, PixelFormat.Format32bppPArgb, PixelFormat.Format32bppRgb
                BIH.biBitCount = 32
                PalEntries = 0
        End Select
        BIH.biClrUsed = PalEntries

        MMFsize = Marshal.SizeOf(BIH) + PalEntries * Marshal.SizeOf(GetType(RGBQUAD)) + ImageSize

        MMFhandle = CreateFileMappingMy(&HFFFFFFFF, 0&, PAGE_READWRITE, 0, MMFsize, 0&)
        If Not MMFhandle.Equals(IntPtr.Zero) Then
            MMFview = MapViewOfFileMy(MMFhandle, FILE_MAP_WRITE, 0, 0, 0)
            If MMFview.Equals(IntPtr.Zero) Then
                CloseHandle(MMFhandle)
            Else
                Marshal.StructureToPtr(BIH, MMFview, True)

                Offset = MMFview.ToInt32 + Marshal.SizeOf(BIH)
                For PalEntry = 0 To PalEntries - 1
                    rgb.rgbRed = BMP.Palette.Entries(PalEntry).R
                    rgb.rgbGreen = BMP.Palette.Entries(PalEntry).G
                    rgb.rgbBlue = BMP.Palette.Entries(PalEntry).B
                    rgbGC = GCHandle.Alloc(rgb, GCHandleType.Pinned)
                    CopyMemory(Offset, rgbGC.AddrOfPinnedObject, Marshal.SizeOf(rgb))
                    rgbGC.Free()
                    Offset = Offset + Marshal.SizeOf(rgb)
                Next
                CopyMemory(Offset, BMPData.Scan0, ImageSize)

                UnmapViewOfFileMy(MMFview)
                ConvertBitmapToMMF2 = MMFhandle
            End If
        End If
        BMP.UnlockBits(BMPData)
        BMPData = Nothing
        BMP.Dispose()
        BMP = Nothing

        If MMFhandle.Equals(IntPtr.Zero) Then
            MsgBox("Failed to convert bitmap", MsgBoxStyle.Critical, "ConvertBitmapToMMF2")
        End If

    End Function

    ' Convert a global memory block to a bitmap
    Private Function ConvertMemoryBlockToBitmap(ByVal hMem As IntPtr) As Bitmap
        Dim BMP As Bitmap
        Dim BIH As BITMAPINFOHEADER
        Dim bihPtr As IntPtr
        Dim dataPtr As IntPtr
        Dim palPtr As IntPtr
        Dim HdrSize As Integer
        Dim PixFormat As PixelFormat
        Dim PalEntries As Integer
        Dim rgb As RGBQUAD

        BMP = Nothing
        bihPtr = GlobalLock(hMem)

        If Not bihPtr.Equals(IntPtr.Zero) Then

            BIH = CType(Marshal.PtrToStructure(bihPtr, GetType(BITMAPINFOHEADER)), BITMAPINFOHEADER)
            HdrSize = BIH.biSize
            palPtr = New IntPtr(bihPtr.ToInt32() + HdrSize)

            ' Most of these formats are untested
            PixFormat = PixelFormat.Format1bppIndexed
            Select Case BIH.biBitCount
                Case 1
                    HdrSize += 2 * Marshal.SizeOf(rgb)
                    PixFormat = PixelFormat.Format1bppIndexed
                    PalEntries = 2
                Case 4
                    HdrSize += 16 * Marshal.SizeOf(rgb)
                    PixFormat = PixelFormat.Format4bppIndexed
                    PalEntries = BIH.biClrUsed
                Case 8
                    HdrSize += 256 * Marshal.SizeOf(rgb)
                    PixFormat = PixelFormat.Format8bppIndexed
                    PalEntries = BIH.biClrUsed
                Case 16
                    ' Account for the 3 DWORD colour mask
                    If BIH.biCompression = BI_BITFIELDS Then HdrSize += 12
                    PixFormat = PixelFormat.Format16bppRgb555
                    PalEntries = 0
                Case 24
                    PixFormat = PixelFormat.Format24bppRgb
                    PalEntries = 0
                Case 32
                    ' Account for the 3 DWORD colour mask
                    If BIH.biCompression = BI_BITFIELDS Then HdrSize += 12
                    PixFormat = PixelFormat.Format32bppRgb
                    PalEntries = 0
                Case Else
            End Select

            dataPtr = New IntPtr(bihPtr.ToInt32() + HdrSize)
            BMP = New Bitmap(BIH.biWidth, Math.Abs(BIH.biHeight), PixFormat)
            If PalEntries > 0 Then

                palPtr = New IntPtr(bihPtr.ToInt32() + BIH.biSize)
                Dim Pal As ColorPalette
                Dim PalEntry As Integer
                Pal = BMP.Palette
                For PalEntry = 0 To PalEntries - 1
                    rgb = CType(Marshal.PtrToStructure(palPtr, GetType(RGBQUAD)), RGBQUAD)
                    Pal.Entries(PalEntry) = Color.FromArgb(rgb.rgbRed, rgb.rgbGreen, rgb.rgbBlue)
                    palPtr = New IntPtr(palPtr.ToInt32() + Marshal.SizeOf(rgb))
                Next PalEntry
                BMP.Palette = Pal
            End If
            Dim BMPData As BitmapData
            BMPData = BMP.LockBits(New Rectangle(New Point, BMP.Size), ImageLockMode.ReadWrite, PixFormat)
            CopyMemory(BMPData.Scan0.ToInt32(), dataPtr, BMPData.Stride * BMP.Height)
            BMP.UnlockBits(BMPData)
            ' Flip the bitmap (GDI+ bitmap scan lines are top down, GDI are bottom up)
            BMP.RotateFlip(RotateFlipType.RotateNoneFlipY)

            ' Reset the resolutions
            BMP.SetResolution(CSng(Int(BIH.biXPelsPerMeter * 2.54 / 100 + 0.5)), CSng(Int(BIH.biYPelsPerMeter * 2.54 / 100 + 0.5)))
            GlobalUnlock(hMem)
        End If

        Return BMP
    End Function

End Module
