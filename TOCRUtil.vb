'
' THE SOFTWARE IS PROVIDED "AS-IS" AND WITHOUT WARRANTY OF ANY KIND, 
' EXPRESS, IMPLIED OR OTHERWISE, INCLUDING WITHOUT LIMITATION, ANY 
' WARRANTY OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.  
'
' Copyright (C) 2022 Transym Computer Services Ltd.
'
' TOCR5DemoVB.NET

Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices
Imports System.Drawing.Imaging
Module TOCRUtil
#Region " SDK Declares "
    Private Const DIB_RGB_COLORS As Integer = 0
    Private Const BI_RGB As Integer = 0
    Private Const BI_BITFIELDS As Integer = 3
    Private Const PAGE_READWRITE As Integer = 4
    Private Const FILE_MAP_WRITE As Integer = 2
    Private Const SRCCOPY As Integer = &HCC0020&

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Structure RGBQUAD
        Dim rgbBlue As Byte
        Dim rgbGreen As Byte
        Dim rgbRed As Byte
        Dim rgbReserved As Byte
    End Structure ' RGBQUAD

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
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


    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Structure BITMAPINFO
        Dim bmih As BITMAPINFOHEADER
        <VBFixedArray(2), MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)>
        Public cols As UInt32()
    End Structure ' BITMAPINFO

    Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Boolean
    Private Declare Function CreateFileMappingMy Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Integer, ByVal lpFileMappigAttributes As Integer, ByVal flProtect As Integer, ByVal dwMaximumSizeHigh As Integer, ByVal dwMaximumSizeLow As Integer, ByVal lpName As Integer) As IntPtr
    Private Declare Function MapViewOfFileMy Lib "kernel32" Alias "MapViewOfFile" (ByVal hFileMappingObject As IntPtr, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Integer, ByVal dwFileOffsetLow As Integer, ByVal dwNumberOfBytesToMap As Integer) As IntPtr
    Private Declare Function UnmapViewOfFileMy Lib "kernel32" Alias "UnmapViewOfFile" (ByVal lpBaseAddress As IntPtr) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As Integer, ByVal lpvSrc As IntPtr, ByVal cbCopy As Integer)
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As IntPtr) As IntPtr
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As IntPtr) As Integer
    Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As IntPtr) As Integer

    Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As IntPtr) As Boolean
    Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As IntPtr) As IntPtr
    Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hRefDC As IntPtr) As IntPtr

    Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Boolean
    Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As IntPtr) As Boolean
    Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As IntPtr, ByVal hObject As IntPtr) As IntPtr
    Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hdc As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As Integer) As Boolean
    Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As IntPtr, ByRef pbmi As BITMAPINFO, ByVal iUsage As UInt32, ByRef ppvBits As IntPtr, ByVal hSection As IntPtr, ByVal dwOffset As UInt32) As IntPtr

#End Region

#Region " Bitmap Conversion "

    ' Convert a bitmap to 1bpp' Transym OCR Demonstration program
    Public Function ConvertTo1bpp(ByVal BMPIn As Bitmap) As Bitmap

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
    Public Function ConvertBitmapToMMF(ByVal BMPIn As Bitmap,
        Optional ByVal DiscardBitmap As Boolean = True,
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
    Public Function ConvertBitmapToMMF2(ByRef BMPIn As Bitmap,
        Optional ByVal DiscardBitmap As Boolean = True,
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
    Public Function ConvertMemoryBlockToBitmap(ByVal hMem As IntPtr) As Bitmap
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

#End Region
End Module
