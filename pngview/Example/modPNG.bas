Attribute VB_Name = "modPNG"
Option Explicit

'--------- Helper functions from the VBPNG library
' These are just wrappers around the C runtime calls, so VB can call them
Private Declare Function crt_fopen Lib "vbpng1.dll" (ByVal filename _
    As String, ByVal mode As String) As Long
Private Declare Function crt_fclose Lib "vbpng1.dll" (ByVal stream _
    As Long) As Long
Private Declare Function crt_fread Lib "vbpng1.dll" (buffer As Any, _
    ByVal nsize As Long, ByVal nCount As Long, ByVal stream As _
    Long) As Long
Private Declare Function crt_malloc Lib "vbpng1.dll" (ByVal nsize _
    As Long) As Long
Private Declare Sub crt_free Lib "vbpng1.dll" (ByVal lpVoid As Long)

'--------- PNG API's
Private Type png_color_8
   red As Byte
   green As Byte
   blue As Byte
   gray As Byte
   alpha As Byte
End Type

Private Declare Function png_sig_cmp Lib "vbpng1.dll" (sig As Any, _
    ByVal start As Long, ByVal numtocheck As Long) As Long
Private Declare Function png_create_read_struct Lib "vbpng1.dll" _
    (ByVal user_png_ver As String, ByVal error_ptr As Long, _
    ByVal error_fn As Long, ByVal warn_fn As Long) As Long
Private Declare Function png_create_info_struct Lib "vbpng1.dll" _
    (ByVal png_ptr As Long) As Long
Private Declare Sub png_init_io Lib "vbpng1.dll" (ByVal png_ptr _
    As Long, ByVal fp As Long)
Private Declare Sub png_set_sig_bytes Lib "vbpng1.dll" (ByVal _
    png_ptr As Long, ByVal num As Long)
Private Declare Sub png_read_png Lib "vbpng1.dll" (ByVal png_ptr _
    As Long, ByVal info_ptr As Long, ByVal transforms As Long, _
    ByVal params As Long)
Private Declare Sub png_read_end Lib "vbpng1.dll" (ByVal png_ptr _
    As Long, ByVal end_info As Long)
Private Declare Sub png_destroy_read_struct Lib "vbpng1.dll" _
    (png_ptr As Long, info_ptr As Long, end_info As Long)
Private Declare Function png_get_rows Lib "vbpng1.dll" (ByVal _
    png_ptr As Long, ByVal info_ptr As Long) As Long
Private Declare Function png_get_IHDR Lib "vbpng1.dll" (ByVal _
    png_ptr As Long, ByVal info_ptr As Long, Width As Long, _
    Height As Long, bit_depth As Long, color_type As Long, _
    interlace_type As Long, compression_type As Long, _
    filter_type As Long) As Long
Private Declare Sub png_read_image Lib "vbpng1.dll" (ByVal _
    png_ptr As Long, ByVal row_pointers As Long)
Private Declare Function png_get_rowbytes Lib "vbpng1.dll" _
    (ByVal png_ptr As Long, ByVal info_ptr As Long) As Long
Private Declare Sub png_read_info Lib "vbpng1.dll" (ByVal _
    png_ptr As Long, ByVal info_ptr As Long)
Private Declare Function png_ws_getrow Lib "vbpng1.dll" (ByVal _
    info_ptr As Long, ByVal nRow As Long) As Long
Private Declare Function png_create_write_struct Lib "vbpng1.dll" _
    (ByVal verstring As String, ByVal user_error_ptr As Long, _
    ByVal user_error_fn As Long, ByVal user_warning_fn As Long) _
    As Long
Private Declare Sub png_set_IHDR Lib "vbpng1.dll" (ByVal png_ptr _
    As Long, ByVal info_ptr As Long, ByVal Width As Long, ByVal _
    Height As Long, ByVal bit_depth As Long, ByVal color_type As _
    Long, ByVal interlace_type As Long, ByVal compression_type _
    As Long, ByVal filter_type As Long)
Private Declare Sub png_set_sBIT Lib "vbpng1.dll" (ByVal png_ptr _
    As Long, ByVal info_ptr As Long, sig_bit As png_color_8)
Private Declare Sub png_write_info Lib "vbpng1.dll" (ByVal _
    png_ptr As Long, ByVal info_ptr As Long)
Private Declare Sub png_set_bgr Lib "vbpng1.dll" (ByVal _
    png_ptr As Long)
Private Declare Sub png_write_image Lib "vbpng1.dll" (ByVal _
    png_ptr As Long, row_pointers As Any)
Private Declare Sub png_write_end Lib "vbpng1.dll" (ByVal _
    png_ptr As Long, ByVal info_ptr As Long)
Private Declare Sub png_destroy_write_struct Lib "vbpng1.dll" _
    (png_ptr As Long, info_ptr As Long)

'Private Declare Function z Lib "vbpng1.dll" zz

Private Const PNG_COLOR_MASK_COLOR = 2
Private Const PNG_COLOR_TYPE_RGB = (PNG_COLOR_MASK_COLOR)
Private Const PNG_INTERLACE_NONE = 0
Private Const PNG_COMPRESSION_TYPE_BASE = 0
Private Const PNG_FILTER_TYPE_BASE = 0

Private Const PNG_TRANSFORM_IDENTITY = &H0&
Private Const PNG_TRANSFORM_STRIP_16 = &H1&
Private Const PNG_TRANSFORM_STRIP_ALPHA = &H2&
Private Const PNG_TRANSFORM_PACKING = &H4&
Private Const PNG_TRANSFORM_PACKSWAP = &H8&
Private Const PNG_TRANSFORM_EXPAND = &H10&
Private Const PNG_TRANSFORM_INVERT_MONO = &H20&
Private Const PNG_TRANSFORM_SHIFT = &H40&
Private Const PNG_TRANSFORM_BGR = &H80&
Private Const PNG_TRANSFORM_SWAP_ALPHA = &H100&
Private Const PNG_TRANSFORM_SWAP_ENDIAN = &H200&
Private Const PNG_TRANSFORM_INVERT_ALPHA = &H400&
Private Const PNG_TRANSFORM_STRIP_FILLER = &H800&

Private Const PNG_LIBPNG_VER_STRING = "1.0.8"

'--------- Standard Windows API's

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias _
    "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" _
    (ByVal hLibModule As Long) As Long

'Save a PictureBox's content into a file
Public Function SavePNGFile(pct As PictureBox, _
    sPNGFileName As String) As Boolean
    
    InIDEHack
    
    Dim fp As Long
    Dim png_ptr As Long
    Dim info_ptr As Long
    Dim nWidth As Long
    Dim nHeight As Long
    Dim sig_bit As png_color_8
    Dim row_pointers() As Long
    Dim i As Long
    Dim dib As cDIBSection
    Dim nPointer As Long
    Dim nPerRow As Long
    
    nWidth = pct.ScaleX(pct.Width, pct.ScaleMode, vbPixels)
    nHeight = pct.ScaleY(pct.Height, pct.ScaleMode, vbPixels)
    
    If nWidth <= 0 Or nHeight <= 0 Then
        Exit Function
    End If
    
    fp = crt_fopen(sPNGFileName, "wb")
    If fp = 0 Then
        Exit Function
    End If

    png_ptr = png_create_write_struct(PNG_LIBPNG_VER_STRING, _
        0, 0, 0)
    
    If png_ptr = 0 Then
        crt_fclose fp
        Exit Function
    End If

    info_ptr = png_create_info_struct(png_ptr)
    If info_ptr = 0 Then
        crt_fclose fp
        Exit Function
    End If

    png_init_io png_ptr, fp

    png_set_IHDR png_ptr, info_ptr, nWidth, nHeight, 8, _
        PNG_COLOR_TYPE_RGB, PNG_INTERLACE_NONE, _
        PNG_COMPRESSION_TYPE_BASE, PNG_FILTER_TYPE_BASE

    sig_bit.gray = 0
    sig_bit.red = 8
    sig_bit.green = 8
    sig_bit.blue = 8
    sig_bit.alpha = 0
    
    png_set_sBIT png_ptr, info_ptr, sig_bit

    png_write_info png_ptr, info_ptr

    png_set_bgr png_ptr
    
    ReDim row_pointers(nHeight - 1)
    Set dib = New cDIBSection
    dib.Create nWidth, nHeight
    dib.LoadPictureBlt pct.hdc
    
    nPerRow = dib.BytesPerScanLine
    nPointer = dib.DIBSectionBitsPtr
    
    For i = 0 To nHeight - 1
        row_pointers(nHeight - i - 1) = nPointer
        nPointer = nPointer + nPerRow
    Next
    
    png_write_image png_ptr, row_pointers(0)
    
    png_write_end png_ptr, info_ptr
    png_destroy_write_struct png_ptr, info_ptr
    crt_fclose fp
    dib.ClearUp
    Set dib = Nothing
    
    SavePNGFile = True
End Function

'Load a PNG file into a PictureBox
Public Function LoadPNGFile(sPNGFileName As String, _
    pct As PictureBox) As Boolean
    
    InIDEHack

    Dim fp As Long
    Dim header(7) As Byte
    Dim png_ptr As Long
    Dim info_ptr As Long
    Dim end_info As Long
    Dim nWidth As Long
    Dim nHeight As Long
    Dim nTemp(4) As Long
    Dim dib As cDIBSection
    Dim nPointerDib As Long
    Dim nPerRow As Long
    Dim nPerRowDIB As Long
    Dim i As Long
    Dim bytes() As Byte

    fp = crt_fopen(sPNGFileName, "rb")
    If fp = 0 Then
        Exit Function
    End If
    
    crt_fread header(0), 1, 8, fp
    If png_sig_cmp(header(0), 0, 8) <> 0 Then
        crt_fclose fp
        Exit Function
    End If
    
    png_ptr = png_create_read_struct(PNG_LIBPNG_VER_STRING, _
        0, 0, 0)
        
    If png_ptr = 0 Then
        crt_fclose fp
        Exit Function
    End If
    
    info_ptr = png_create_info_struct(png_ptr)
    end_info = png_create_info_struct(png_ptr)
    If info_ptr = 0 Or end_info = 0 Then
        crt_fclose fp
        Exit Function
    End If
    
    png_init_io png_ptr, fp
    png_set_sig_bytes png_ptr, 8
    
    png_read_png png_ptr, info_ptr, PNG_TRANSFORM_BGR Or _
        PNG_TRANSFORM_STRIP_ALPHA Or PNG_TRANSFORM_PACKING, 0
    png_get_IHDR png_ptr, info_ptr, nWidth, nHeight, nTemp(0), _
        nTemp(1), nTemp(2), nTemp(3), nTemp(4)
    
    If nHeight <= 0 Then
        crt_fclose fp
        Exit Function
    End If
    
    Set dib = New cDIBSection
    dib.Create nWidth, nHeight
    
    nPerRowDIB = dib.BytesPerScanLine
    nPerRow = png_get_rowbytes(png_ptr, info_ptr)
    nPointerDib = dib.DIBSectionBitsPtr
    
    ReDim bytes(nPerRowDIB)
    For i = 0 To nPerRowDIB - 1
        bytes(i) = Int(Rnd * 256)
    Next
    For i = 0 To nHeight - 1
        CopyMemory ByVal nPointerDib, ByVal _
            png_ws_getrow(info_ptr, nHeight - i - 1), nPerRow
        nPointerDib = nPointerDib + nPerRowDIB
    Next
    
    nPointerDib = dib.DIBSectionBitsPtr
    
    pct.Width = pct.ScaleX(nWidth, vbPixels, pct.ScaleMode)
    pct.Height = pct.ScaleY(nHeight, vbPixels, pct.ScaleMode)
    pct.Cls
    dib.PaintPicture pct.hdc, , , nWidth, nHeight
    pct.Refresh
        
    png_destroy_read_struct png_ptr, info_ptr, end_info
    crt_fclose fp
    dib.ClearUp
    Set dib = Nothing
    
    LoadPNGFile = True
    
End Function

Private Sub InIDEHack()
'Small hack to manually load the DLL if we're in the IDE
' since the IDE doesn't look in App.Path for DLLs, where as
' the executable will
On Error GoTo InIDE
    'This will only raise an error in the IDE
    Debug.Assert 1 / 0

    Exit Sub
InIDE:
    Dim hLib As Long
    'Since where in the IDE, load the DLL
    hLib = LoadLibrary(App.Path & "\vbpng1.dll")
    'Call a useless function to cause VB to load the DLL
    crt_free 0
    'Free our reference to the DLL, now VB has one
    FreeLibrary hLib
End Sub

