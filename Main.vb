' Transym OCR Demonstration program
'
' THE SOFTWARE IS PROVIDED "AS-IS" AND WITHOUT WARRANTY OF ANY KIND, 
' EXPRESS, IMPLIED OR OTHERWISE, INCLUDING WITHOUT LIMITATION, ANY 
' WARRANTY OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.  
'
' This program demonstrates calling TOCR version 5.1 from VB.NET.
'
' Copyright (C) 2022 Transym Computer Services Ltd.
'
'
' TOCR5 Demo VB.NET

Option Strict On
Option Explicit On 

Imports System.Runtime.InteropServices
Imports System.Drawing.Imaging

Module Main
#If PLATFORM = "x64" Then
    Private mSample_TIF_file As String = Application.StartupPath & "\..\..\..\Sample.tif"
    Private mSample_BMP_file As String = Application.StartupPath & "\..\..\..\Sample.bmp"
    Private mSample_PDF_file As String = Application.StartupPath & "\..\..\..\Sample.pdf"
#Else
    Private mSample_TIF_file As String = Application.StartupPath & "\..\..\Sample.tif"
    Private mSample_BMP_file As String = Application.StartupPath & "\..\..\Sample.bmp"
    Private mSample_PDF_file As String = Application.StartupPath & "\..\..\Sample.pdf"
#End If

    Sub Main()
        Example1() ' Demonstrates how to OCR a file
        Example2() ' Demonstrates how to OCR multiple files
        Example3() ' Demonstrates how to OCR an image using a memory mapped file created by TOCR & Demonstrates TOCRRESULTS_EG
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
        Dim JobInfo_EG As New TOCRJOBINFO_EG
        Dim Answer As String = ""
        Dim Results As New TOCRRESULTS_EG

        TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX)

        JobInfo_EG.Initialize()

        JobInfo_EG.InputFile = mSample_TIF_file
        JobInfo_EG.JobType = TOCRJOBTYPE_TIFFFILE
        ' or
        'JobInfo_EG.InputFile = mSample_BMP_file
        'JobInfo_EG.JobType = TOCRJOBTYPE_DIBFILE
        ' or
        'JobInfo_EG.InputFile = mSample_PDF_file
        'JobInfo_EG.JobType = TOCRJOBTYPE_PDFFILE

        Status = TOCRInitialise(JobNo)
        If Status = TOCR_OK Then
            If OCRWait(JobNo, JobInfo_EG) Then
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
        Dim JobInfo_EG As New TOCRJOBINFO_EG
        Dim Results As New TOCRRESULTS_EG
        Dim CountDone As Integer = 0

        TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX)

        JobInfo_EG.Initialize()

        Status = TOCRInitialise(JobNo)
        If Status = TOCR_OK Then

            ' 1st file
            JobInfo_EG.InputFile = mSample_TIF_file
            JobInfo_EG.JobType = TOCRJOBTYPE_TIFFFILE
            If OCRWait(JobNo, JobInfo_EG) Then
                If GetResults(JobNo, Results) Then
                    CountDone += 1
                End If
            End If

            ' 2nd file
            JobInfo_EG.InputFile = mSample_BMP_file
            JobInfo_EG.JobType = TOCRJOBTYPE_DIBFILE
            If OCRWait(JobNo, JobInfo_EG) Then
                If GetResults(JobNo, Results) Then
                    CountDone += 1
                End If
            End If

            ' 3rd file
            JobInfo_EG.InputFile = mSample_PDF_file
            JobInfo_EG.JobType = TOCRJOBTYPE_PDFFILE
            If OCRWait(JobNo, JobInfo_EG) Then
                If GetResults(JobNo, Results) Then
                    CountDone += 1
                End If
            End If
            TOCRShutdown(JobNo)
        End If

        MsgBox(CountDone.ToString() & " of 3 jobs done", MsgBoxStyle.Information, "Example 2")
    End Sub ' Example 2

    ' Demonstrates how to OCR an image using a memory mapped file created by TOCR & Demonstrates TOCRRESULTS_EG
    Private Sub Example3()

        Dim Status As Integer
        Dim JobNo As Integer
        Dim Answer As String = ""
        Dim MMFhandle As IntPtr = IntPtr.Zero
        Dim Results As New TOCRRESULTSEX_EG
        Dim JobInfo_EG As New TOCRJOBINFO_EG

        TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX)

        JobInfo_EG.Initialize()

        JobInfo_EG.JobType = TOCRJOBTYPE_MMFILEHANDLE

        Status = TOCRInitialise(JobNo)
        If Status = TOCR_OK Then
            Status = TOCRConvertFormat(JobNo, mSample_TIF_file, TOCRCONVERTFORMAT_TIFFFILE, MMFhandle, TOCRCONVERTFORMAT_MMFILEHANDLE, 0)
            If Status = TOCR_OK Then
                JobInfo_EG.hMMF = MMFhandle
                If OCRWait(JobNo, JobInfo_EG) Then
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
        Dim JobInfo_EG As New TOCRJOBINFO_EG
        Dim Answer As String = ""
        Dim MMFhandle As IntPtr = IntPtr.Zero
        Dim Results As New TOCRRESULTS_EG

        BMP = New Bitmap(mSample_BMP_file)

        MMFhandle = ConvertBitmapToMMF(BMP)
        'MMFhandle = ConvertBitmapToMMF2(BMP)

        If Not MMFhandle.Equals(IntPtr.Zero) Then

            TOCRSetConfig(TOCRCONFIG_DEFAULTJOB, TOCRCONFIG_DLL_ERRORMODE, TOCRERRORMODE_MSGBOX)

            JobInfo_EG.Initialize()

            JobInfo_EG.JobType = TOCRJOBTYPE_MMFILEHANDLE

            Status = TOCRInitialise(JobNo)

            If Status = TOCR_OK Then
                JobInfo_EG.hMMF = MMFhandle

                If OCRWait(JobNo, JobInfo_EG) Then
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
    Private Function OCRWait(ByVal JobNo As Integer, ByVal JobInfo_EG As TOCRJOBINFO_EG) As Boolean

        Dim Status As Integer
        Dim JobStatus As Integer
        Dim Msg As String
        Dim ErrorMode As Integer


        Status = TOCRDoJob_EG(JobNo, JobInfo_EG)
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
    Private Function OCRPoll(ByVal JobNo As Integer, ByVal JobInfo_EG As TOCRJOBINFO_EG) As Boolean

        Dim Status As Integer
        Dim JobStatus As Integer
        Dim Msg As String
        Dim ErrorMode As Integer
        Dim Progress As Single
        Dim AutoOrientation As Integer

        Status = TOCRDoJob_EG(JobNo, JobInfo_EG)
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
    Private Function GetResults(ByVal JobNo As Integer, ByRef Results As TOCRRESULTS_EG) As Boolean

        Dim ResultsInf As Integer ' number of bytes needed for results
        Dim BytesGC As GCHandle
        Dim AddrOfItemBytes As System.IntPtr
        Dim ItemNo As Integer
        Dim Bytes() As Byte
        Dim Offset As Integer

        GetResults = False
        Results.Hdr.NumItems = 0

        If TOCRGetJobResultsEx_EG(JobNo, TOCRGETRESULTS_NORMAL_EG, ResultsInf, IntPtr.Zero) = TOCR_OK Then
            If ResultsInf > 0 Then
                ReDim Bytes(ResultsInf - 1)
                Dim s As Integer
                For s = 0 To ResultsInf - 1 : Bytes(s) = 255 : Next s

                ' pin the Bytes array so that TOCRGetJobResults can write to it
                BytesGC = GCHandle.Alloc(Bytes, GCHandleType.Pinned)

                If TOCRGetJobResultsEx_EG(JobNo, TOCRGETRESULTS_NORMAL_EG, ResultsInf, BytesGC.AddrOfPinnedObject) = TOCR_OK Then
                    With Results
                        .Hdr = CType(Marshal.PtrToStructure(BytesGC.AddrOfPinnedObject, GetType(TOCRRESULTSHEADER_EG)), TOCRRESULTSHEADER_EG)
                        If .Hdr.NumItems > 0 Then
                            ReDim .Item(.Hdr.NumItems - 1)
                            Offset = Marshal.SizeOf(GetType(TOCRRESULTSHEADER_EG))
                            For ItemNo = 0 To .Hdr.NumItems - 1
                                AddrOfItemBytes = Marshal.UnsafeAddrOfPinnedArrayElement(Bytes, Offset)
                                .Item(ItemNo) = CType(Marshal.PtrToStructure(AddrOfItemBytes, GetType(TOCRRESULTSITEM_EG)), TOCRRESULTSITEM_EG)
                                Offset = Offset + Marshal.SizeOf(GetType(TOCRRESULTSITEM_EG))
                            Next ItemNo
                        End If ' .Hdr.NumItems > 0

                        GetResults = True

                    End With ' results
                End If ' TOCRGetJobResults(JobNo, ResultsInf, Bytes(0)) = TOCR_OK

                BytesGC.Free()

            End If ' ResultsInf > 0
        End If ' TOCRGetJobResults(JobNo, ResultsInf, 0) = TOCR_OK

    End Function

    ' copy of TOCRRESULTSITEMEX_EG without the Alt[] array 
    Private Structure TOCRRESULTSITEMEXHDR_EG
        Dim Confidence As Single
        Dim StructId As Short
        Dim OCRCharWUnicode As Short 'V5 split from OCRChaW
        Dim OCRCharWInternal As Short 'V5 split from OCRChaW
        Dim FontID As Short
        Dim FontStyleInfo As Short
        Dim XPos As Short
        Dim YPos As Short
        Dim XDim As Short
        Dim YDim As Short
        Dim YDimRef As Short
        Dim Noise As Short 'V5 addition
    End Structure

    ' OVERLOADED function to retrieve the extended results from the service process and load into 'Results'
    Function GetResults(ByVal JobNo As Integer, ByRef Results As TOCRRESULTSEX_EG) As Boolean
        Dim ResultsInf As Integer ' number of bytes needed for results
        Dim AddrOfItemBytes As System.IntPtr
        Dim ItemNo As Integer           ' loop counter
        Dim AltNo As Integer            ' loop counter
        Dim Bytes() As Byte             ' array of bytes of returned results
        Dim BytesGC As GCHandle         ' handle ti pin Bytes()
        Dim Offset As Integer           ' address offset into Bytes()
        Dim ItemHdr As TOCRRESULTSITEMEXHDR_EG

        GetResults = False
        Results.Hdr.NumItems = 0
        If TOCRGetJobResultsEx_EG(JobNo, TOCRGETRESULTS_EXTENDED_EG, ResultsInf, IntPtr.Zero) = TOCR_OK Then
            If ResultsInf > 0 Then
                ReDim Bytes(ResultsInf - 1)
                ' pin the Bytes array so that TOCRGetJobResultsEx can write to it
                BytesGC = GCHandle.Alloc(Bytes, GCHandleType.Pinned)

                If TOCRGetJobResultsEx_EG(JobNo, TOCRGETRESULTS_EXTENDED_EG, ResultsInf, BytesGC.AddrOfPinnedObject) = TOCR_OK Then
                    With Results
                        .Hdr = CType(Marshal.PtrToStructure(BytesGC.AddrOfPinnedObject, GetType(TOCRRESULTSHEADER_EG)), TOCRRESULTSHEADER_EG)
                        If .Hdr.NumItems > 0 Then
                            ReDim .Item(.Hdr.NumItems - 1)
                            Offset = Marshal.SizeOf(GetType(TOCRRESULTSHEADER_EG))
                            For ItemNo = 0 To .Hdr.NumItems - 1
                                AddrOfItemBytes = Marshal.UnsafeAddrOfPinnedArrayElement(Bytes, Offset)
                                ' Cannot Marshal TOCRRESULTSITEMEX_EG so use copy of structure header
                                ' This unfortunately means a double copy of the data
                                ItemHdr = CType(Marshal.PtrToStructure(AddrOfItemBytes, GetType(TOCRRESULTSITEMEXHDR_EG)), TOCRRESULTSITEMEXHDR_EG)
                                With .Item(ItemNo)
                                    .Initialize()
                                    .Confidence = ItemHdr.Confidence
                                    .StructId = ItemHdr.StructId
                                    .OCRCharWUnicode = ItemHdr.OCRCharWUnicode
                                    .OCRCharWInternal = ItemHdr.OCRCharWInternal
                                    .FontID = ItemHdr.FontID
                                    .FontStyleInfo = ItemHdr.FontStyleInfo
                                    .XPos = ItemHdr.XPos
                                    .YPos = ItemHdr.YPos
                                    .XDim = ItemHdr.XDim
                                    .YDim = ItemHdr.YDim
                                    .YDimRef = ItemHdr.YDimRef
                                    .Noise = ItemHdr.Noise
                                    Offset = Offset + Marshal.SizeOf(GetType(TOCRRESULTSITEMEXHDR_EG))
                                    For AltNo = 0 To 4
                                        AddrOfItemBytes = Marshal.UnsafeAddrOfPinnedArrayElement(Bytes, Offset)
                                        .Alt(AltNo) = CType(Marshal.PtrToStructure(AddrOfItemBytes, GetType(TOCRRESULTSITEMEXALT_EG)), TOCRRESULTSITEMEXALT_EG)
                                        Offset = Offset + Marshal.SizeOf(GetType(TOCRRESULTSITEMEXALT_EG))
                                    Next AltNo
                                End With
                            Next ItemNo
                        End If ' .Hdr.NumItems > 0

                        GetResults = True

                    End With ' results
                End If ' TOCRGetJobResults_EG(JobNo, ResultsInf, Bytes(0)) = TOCR_OK

                BytesGC.Free()

            End If ' ResultsInf > 0
        End If ' TOCRGetJobResults_EG(JobNo, ResultsInf, 0) = TOCR_OK

    End Function

    'OVERLOADED function to convert results to a string
    Private Function FormatResults(ByVal Results As TOCRRESULTS_EG, ByRef Answer As String) As Boolean

        Dim ItemNo As Integer

        FormatResults = False
        Answer = ""

        With Results
            If .Hdr.NumItems > 0 Then
                For ItemNo = 0 To .Hdr.NumItems - 1
                    If ChrW(.Item(ItemNo).OCRCharWUnicode) = vbCr Then
                        Answer = Answer & vbCrLf
                    Else
                        Answer = Answer & ChrW(.Item(ItemNo).OCRCharWUnicode)
                    End If
                Next ItemNo
                FormatResults = True
            Else
                MsgBox("No results returned", MsgBoxStyle.Information, "FormatResults")
            End If
        End With

    End Function

    'OVERLOADED function to convert extended results to a string
    Private Function FormatResults(ByVal Results As TOCRRESULTSEX_EG, ByRef Answer As String) As Boolean

        Dim ItemNo As Integer

        FormatResults = False

        With Results
            If .Hdr.NumItems > 0 Then
                For ItemNo = 0 To .Hdr.NumItems - 1
                    If ChrW(.Item(ItemNo).OCRCharWUnicode) = vbCr Then
                        Answer = Answer & vbCrLf
                    Else
                        Answer = Answer & ChrW(.Item(ItemNo).OCRCharWUnicode)
                    End If
                Next ItemNo
                FormatResults = True
            Else
                MsgBox("No results returned", MsgBoxStyle.Information, "FormatResults")
            End If
        End With

    End Function

End Module
