Attribute VB_Name = "modGradiant"
Option Explicit

Private Type TRIVERTEX
    x           As Long
    y           As Long
    Red         As Integer
    Green       As Integer
    Blue        As Integer
    Alpha       As Integer
End Type


Private Type GRADIENT_RECT
    UpperLeft   As Long
    LowerRight  As Long
End Type


Private Const GRADIENT_FILL_RECT_H      As Long = &H0
Private Const GRADIENT_FILL_RECT_V      As Long = &H1


Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long




Private Function LongToSignedShort(Value As Long) As Integer

   If Value < (2 ^ 16) / 2 Then
      LongToSignedShort = CInt(Value)
   Else
      LongToSignedShort = CInt(Value - 2 ^ 16)
   End If

End Function


Public Sub GradientFill(Dest As PictureBox, Horizontal As Boolean, Colour1 As Long, Colour2 As Long)
    
    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    Dim OldScaleMode As ScaleModeConstants

    OldScaleMode = Dest.ScaleMode
    Dest.ScaleMode = vbPixels

    With vert(0)
        .x = 0
        .y = 0
        .Red = LongToSignedShort(CLng((Colour1 And &HFF&) * 256))
        .Green = LongToSignedShort(CLng(((Colour1 And &HFF00&) \ &H100&) * 256))
        .Blue = LongToSignedShort(CLng(((Colour1 And &HFF0000) \ &H10000) * 256))
        .Alpha = 0
    End With

    With vert(1)
        .x = Dest.ScaleWidth
        .y = Dest.ScaleHeight
        .Red = LongToSignedShort(CLng((Colour2 And &HFF&) * 256))
        .Green = LongToSignedShort(CLng(((Colour2 And &HFF00&) \ &H100&) * 256))
        .Blue = LongToSignedShort(CLng(((Colour2 And &HFF0000) \ &H10000) * 256))
        .Alpha = 0
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    GradientFillRect Dest.hdc, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H
    
    Dest.ScaleMode = OldScaleMode
End Sub

