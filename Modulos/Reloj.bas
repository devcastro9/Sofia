Attribute VB_Name = "Reloj"
Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

'Disable X
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Global Const MF_BYPOSITION = &H400
Global Const MF_REMOVE = &H1000&

Global gbDescarga    As Boolean
Private bFirstTime  As Boolean

Dim iTop0   As Integer
Dim iTop15  As Integer
Dim iTop24  As Integer
Dim iTop3   As Integer
Dim iTop6   As Integer
Dim iTopInv As Integer
Dim iAltForm   As Integer

Sub RLJConfigureForm()
'Purpose: Configure Form and disable X button
'Author: L124RD K1N6

   Dim hSysMenu As Long, nCnt As Long
   
   iTop0 = 300  '0
   iTop15 = 420 '120 '8
   iTop24 = 900 '600  '40
   iTop3 = 1320 '1020 '72
   iTop6 = 780  '480 '36
   iTopInv = 200
   iAltForm = 1600
   
   'hSysMenu = GetSystemMenu(FrmReloj.hwnd, False)
   hSysMenu = GetSystemMenu(Frm_Asistencia.hwnd, False)
   

   If hSysMenu Then
       nCnt = GetMenuItemCount(hSysMenu)
       If nCnt Then
           RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE
           RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE
           DrawMenuBar Frm_Asistencia.hwnd
       End If
   End If

   Call BringWindowToTop(Frm_Asistencia.hwnd)

End Sub

Public Sub RLJGlassifyForm(frm As Form)
'Purpose: Redraw form
'Modified by: L124RD K1N6
'Notes: This code is not mine

    Const RGN_DIFF = 4
    Const RGN_OR = 2
    Dim outer_rgn As Long
    Dim inner_rgn As Long
    Dim wid As Single
    Dim hgt As Single
    Dim border_width As Single
    Dim title_height As Single
    Dim ctl_left As Single
    Dim ctl_top As Single
    Dim ctl_right As Single
    Dim ctl_bottom As Single
    Dim control_rgn As Long
    Dim combined_rgn As Long
    Dim ctl As CONTROL
    
    On Error GoTo ERROR_GLASS
    
    If frm.WindowState = vbMinimized Then Exit Sub
    ' Create the main form region.
    wid = frm.ScaleX(frm.Width, vbTwips, vbPixels)
    hgt = frm.ScaleY(frm.Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)
    border_width = (wid - frm.ScaleWidth) / 2
    title_height = hgt - border_width - frm.ScaleHeight
    inner_rgn = CreateRectRgn(border_width, title_height, wid - border_width, hgt - border_width)
    ' Subtract the inner region from the out
    '     er.
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, _
    inner_rgn, RGN_DIFF
    ' Create the control regions.


    For Each ctl In frm.Controls
      If Not ctl.Name = "TmrReloj" Then
        If ctl.Container Is frm Then
            ctl_left = frm.ScaleX(ctl.Left, frm.ScaleMode, vbPixels) + border_width
            ctl_top = frm.ScaleX(ctl.Top, frm.ScaleMode, vbPixels) + title_height
            ctl_right = frm.ScaleX(ctl.Width, frm.ScaleMode, vbPixels) + ctl_left
            ctl_bottom = frm.ScaleX(ctl.Height, frm.ScaleMode, vbPixels) + ctl_top
            control_rgn = CreateRectRgn(ctl_left, ctl_top, ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, control_rgn, RGN_OR
        End If
      End If
    Next ctl
    
    ' Restrict the window to the region.
    SetWindowRgn frm.hwnd, combined_rgn, True
    
    Exit Sub
    
ERROR_GLASS:
   'Control Error
   Const FILE$ = "\ERROR.LOg"
   
   Open App.Path & FILE For Output As #1
      Print #1, "ERROR: " & Err.Number & " - " & Err.Description
   Close #1
    
End Sub

Sub RLJDisplayTime()
'Purpose: Display time, the secuence of 1 and 0 is the binary number as if you were programming an electronic Display
'Author: L124RD K1N6
'Comments: I excluded the code for seconds (if you make form wider you can see them) because if i
'          showed the paiting would be more dynamic and i got an error.
'          It looks like gdi32.dll doesn't support many calls.
   
   Const ZERO_BIN$ = "1111110"
   Const ONE_BIN$ = "0110000"
   Const TWO_BIN$ = "1101101"
   Const THREE_BIN$ = "1111001"
   Const FOUR_BIN$ = "0110011"
   Const FIVE_BIN$ = "1011011"
   Const SIX_BIN$ = "1011111"
   Const SEVEN_BIN$ = "1110000"
   Const EIGHT_BIN$ = "1111111"
   Const NINE_BIN$ = "1110011"
   
   Const NEGRO = &H0&
   Const VERDE = &HFF00&
   
      
   Dim sHora      As String * 2
   Dim sMinuto    As String * 2
'   Dim sSegundo   As String * 2
   
   Static bFlagDot As Boolean
   
   Static byHor1   As Byte
   Static byHor2   As Byte
   Static byMin1   As Byte
   Static byMin2   As Byte
'   Static bySeg1   As Byte
'   Static bySeg2   As Byte
      
   Dim bFlagHor1  As Boolean
   Dim bFlagHor2  As Boolean
   Dim bFlagMin1  As Boolean
   Dim bFlagMin2  As Boolean
'   Dim bFlagSeg1  As Boolean
'   Dim bFlagSeg2  As Boolean
      
      
   sHora = Format(Hour(Time), "00")
   sMinuto = Format(Minute(Now), "00")
'   sSegundo = Format(Second(Now), "00")
   
   If byHor1 <> Mid$(sHora, 1, 1) Then byHor1 = Mid$(sHora, 1, 1): bFlagHor1 = True
   If byHor2 <> Mid$(sHora, 2, 1) Then byHor2 = Mid$(sHora, 2, 1): bFlagHor2 = True
   
   If byMin1 <> Mid$(sMinuto, 1, 1) Then byMin1 = Mid$(sMinuto, 1, 1): bFlagMin1 = True
   If byMin2 <> Mid$(sMinuto, 2, 1) Then byMin2 = Mid$(sMinuto, 2, 1): bFlagMin2 = True
   
   If bFlagDot Then
      Frm_Asistencia.ShpDot(0).FillColor = NEGRO
      Frm_Asistencia.ShpDot(1).FillColor = NEGRO
   Else
      Frm_Asistencia.ShpDot(0).FillColor = VERDE
      Frm_Asistencia.ShpDot(1).FillColor = VERDE
   End If
   bFlagDot = Not bFlagDot
      
'   If bySeg1 <> Mid$(sSegundo, 1, 1) Then bySeg1 = Mid$(sSegundo, 1, 1): bFlagSeg1 = True
'   If bySeg2 <> Mid$(sSegundo, 2, 1) Then bySeg2 = Mid$(sSegundo, 2, 1): bFlagSeg2 = True
   
   
'   If bFlagSeg2 Then
'      Select Case bySeg2
'         Case 0
'            Call RLJIlumina("Seg2", ZERO_BIN)
'         Case 1
'            Call RLJIlumina("Seg2", ONE_BIN)
'         Case 2
'            Call RLJIlumina("Seg2", TWO_BIN)
'         Case 3
'            Call RLJIlumina("Seg2", THREE_BIN)
'         Case 4
'            Call RLJIlumina("Seg2", CUATRO_BIN)
'         Case 5
'            Call RLJIlumina("Seg2", FIVE_BIN)
'         Case 6
'            Call RLJIlumina("Seg2", SIX_BIN)
'         Case 7
'            Call RLJIlumina("Seg2", SEVEN_BIN)
'         Case 8
'            Call RLJIlumina("Seg2", EIGHT_BIN)
'         Case 9
'            Call RLJIlumina("Seg2", NINE_BIN)
'      End Select
'      RLJGlassifyForm FrmReloj
'   End If
'
'   If bFlagSeg1 Then
'      Select Case bySeg1
'         Case 0
'            Call RLJIlumina("Seg1", ZERO_BIN)
'         Case 1
'            Call RLJIlumina("Seg1", ONE_BIN)
'         Case 2
'            Call RLJIlumina("Seg1", TWO_BIN)
'         Case 3
'            Call RLJIlumina("Seg1", THREE_BIN)
'         Case 4
'            Call RLJIlumina("Seg1", FOUR_BIN)
'         Case 5
'            Call RLJIlumina("Seg1", FIVE_BIN)
'         Case 6
'            Call RLJIlumina("Seg1", SIX_BIN)
'         Case 7
'            Call RLJIlumina("Seg1", SEVEN_BIN)
'         Case 8
'            Call RLJIlumina("Seg1", EIGHT_BIN)
'         Case 9
'            Call RLJIlumina("Seg1", NINE_BIN)
'      End Select
'   End If

   If Not bFirstTime Then    ' Si es la primera vez
      bFlagMin2 = True
      bFlagMin1 = True
      bFlagHor2 = True
      bFlagHor1 = True
      bFirstTime = True
   End If
   
   If bFlagMin2 Then
      Select Case byMin2
         Case 0
            Call RLJIlumina("Min2", ZERO_BIN)
         Case 1
            Call RLJIlumina("Min2", ONE_BIN)
         Case 2
            Call RLJIlumina("Min2", TWO_BIN)
         Case 3
            Call RLJIlumina("Min2", THREE_BIN)
         Case 4
            Call RLJIlumina("Min2", FOUR_BIN)
         Case 5
            Call RLJIlumina("Min2", FIVE_BIN)
         Case 6
            Call RLJIlumina("Min2", SIX_BIN)
         Case 7
            Call RLJIlumina("Min2", SEVEN_BIN)
         Case 8
            Call RLJIlumina("Min2", EIGHT_BIN)
         Case 9
            Call RLJIlumina("Min2", NINE_BIN)
      End Select
      RLJGlassifyForm Frm_Asistencia
   End If
   
   If bFlagMin1 Then
      Select Case byMin1
         Case 0
            Call RLJIlumina("Min1", ZERO_BIN)
         Case 1
            Call RLJIlumina("Min1", ONE_BIN)
         Case 2
            Call RLJIlumina("Min1", TWO_BIN)
         Case 3
            Call RLJIlumina("Min1", THREE_BIN)
         Case 4
            Call RLJIlumina("Min1", FOUR_BIN)
         Case 5
            Call RLJIlumina("Min1", FIVE_BIN)
         Case 6
            Call RLJIlumina("Min1", SIX_BIN)
         Case 7
            Call RLJIlumina("Min1", SEVEN_BIN)
         Case 8
            Call RLJIlumina("Min1", EIGHT_BIN)
         Case 9
            Call RLJIlumina("Min1", NINE_BIN)
      End Select
      Call BringWindowToTop(Frm_Asistencia.hwnd)
      RLJGlassifyForm Frm_Asistencia
   End If
   
   If bFlagHor2 Then
      Select Case byHor2
         Case 0
            Call RLJIlumina("Hor2", ZERO_BIN)
         Case 1
            Call RLJIlumina("Hor2", ONE_BIN)
         Case 2
            Call RLJIlumina("Hor2", TWO_BIN)
         Case 3
            Call RLJIlumina("Hor2", THREE_BIN)
         Case 4
            Call RLJIlumina("Hor2", FOUR_BIN)
         Case 5
            Call RLJIlumina("Hor2", FIVE_BIN)
         Case 6
            Call RLJIlumina("Hor2", SIX_BIN)
         Case 7
            Call RLJIlumina("Hor2", SEVEN_BIN)
         Case 8
            Call RLJIlumina("Hor2", EIGHT_BIN)
         Case 9
            Call RLJIlumina("Hor2", NINE_BIN)
      End Select
      RLJGlassifyForm Frm_Asistencia
   End If
   
   If bFlagHor1 Then
      Select Case byHor1
         Case 0
            Call RLJIlumina("Hor1", ZERO_BIN)
         Case 1
            Call RLJIlumina("Hor1", ONE_BIN)
         Case 2
            Call RLJIlumina("Hor1", TWO_BIN)
         Case 3
            Call RLJIlumina("Hor1", THREE_BIN)
         Case 4
            Call RLJIlumina("Hor1", FOUR_BIN)
         Case 5
            Call RLJIlumina("Hor1", FIVE_BIN)
         Case 6
            Call RLJIlumina("Hor1", SIX_BIN)
         Case 7
            Call RLJIlumina("Hor1", SEVEN_BIN)
         Case 8
            Call RLJIlumina("Hor1", EIGHT_BIN)
         Case 9
            Call RLJIlumina("Hor1", NINE_BIN)
      End Select
      RLJGlassifyForm Frm_Asistencia
   End If
      
   'BringWindowToTop FrmReloj.hwnd
   
End Sub

Sub RLJIlumina(cNombre As String, sSecuencia As String)
'Purpose: Paint display
'Author: L124RD K1N6

   Dim byCont As Byte
   
   With Frm_Asistencia
   
      If cNombre = "Seg2" Then
         
         For byCont = 0 To 6

            If Not Mid$(sSecuencia, byCont + 1, 1) = "0" Then
               If byCont = 0 Then
                  .ImgSeg2(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg2(byCont).Visible = True
                  .ImgSeg2(byCont).Top = iTop0
               ElseIf byCont = 1 Or byCont = 5 Then
                  .ImgSeg2(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg2(byCont).Visible = True
                  .ImgSeg2(byCont).Top = iTop15
               ElseIf byCont = 2 Or byCont = 4 Then
                  .ImgSeg2(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg2(byCont).Visible = True
                  .ImgSeg2(byCont).Top = iTop24
               ElseIf byCont = 3 Then
                  .ImgSeg2(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg2(byCont).Visible = True
                  .ImgSeg2(byCont).Top = iTop3
               ElseIf byCont = 6 Then
                  .ImgSeg2(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg2(byCont).Visible = True
                  .ImgSeg2(byCont).Top = iTop6
               End If
            Else
               .ImgSeg2(byCont).Visible = False
               .ImgSeg2(byCont).Top = iTopInv
            End If

         Next byCont
         
      ElseIf cNombre = "Seg1" Then
         
         For byCont = 0 To 6

            If Not Mid$(sSecuencia, byCont + 1, 1) = "0" Then
               If byCont = 0 Then
                  .ImgSeg1(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg1(byCont).Visible = True
                  .ImgSeg1(byCont).Top = iTop0
               ElseIf byCont = 1 Or byCont = 5 Then
                  .ImgSeg1(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg1(byCont).Visible = True
                  .ImgSeg1(byCont).Top = iTop15
               ElseIf byCont = 2 Or byCont = 4 Then
                  .ImgSeg1(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg1(byCont).Visible = True
                  .ImgSeg1(byCont).Top = iTop24
               ElseIf byCont = 3 Then
                  .ImgSeg1(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg1(byCont).Visible = True
                  .ImgSeg1(byCont).Top = iTop3
               ElseIf byCont = 6 Then
                  .ImgSeg1(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgSeg1(byCont).Visible = True
                  .ImgSeg1(byCont).Top = iTop6
               End If
            Else
               .ImgSeg1(byCont).Visible = False
               .ImgSeg1(byCont).Top = iTopInv
            End If

         Next byCont
      ElseIf cNombre = "Min2" Then
      
         For byCont = 0 To 6

            If Not Mid$(sSecuencia, byCont + 1, 1) = "0" Then
               If byCont = 0 Then
                  .ImgMin2(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin2(byCont).Visible = True
                  .ImgMin2(byCont).Top = iTop0
               ElseIf byCont = 1 Or byCont = 5 Then
                  .ImgMin2(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin2(byCont).Visible = True
                  .ImgMin2(byCont).Top = iTop15
               ElseIf byCont = 2 Or byCont = 4 Then
                  .ImgMin2(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin2(byCont).Visible = True
                  .ImgMin2(byCont).Top = iTop24
               ElseIf byCont = 3 Then
                  .ImgMin2(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin2(byCont).Visible = True
                  .ImgMin2(byCont).Top = iTop3
               ElseIf byCont = 6 Then
                  .ImgMin2(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin2(byCont).Visible = True
                  .ImgMin2(byCont).Top = iTop6
               End If
            Else
               .ImgMin2(byCont).Visible = False
               .ImgMin2(byCont).Top = iTopInv
            End If

         Next byCont
         
      ElseIf cNombre = "Min1" Then
      
         For byCont = 0 To 6
            
            If Not Mid$(sSecuencia, byCont + 1, 1) = "0" Then
               If byCont = 0 Then
                  .ImgMin1(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin1(byCont).Visible = True
                  .ImgMin1(byCont).Top = iTop0
               ElseIf byCont = 1 Or byCont = 5 Then
                  .ImgMin1(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin1(byCont).Visible = True
                  .ImgMin1(byCont).Top = iTop15
               ElseIf byCont = 2 Or byCont = 4 Then
                  .ImgMin1(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin1(byCont).Visible = True
                  .ImgMin1(byCont).Top = iTop24
               ElseIf byCont = 3 Then
                  .ImgMin1(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin1(byCont).Visible = True
                  .ImgMin1(byCont).Top = iTop3
               ElseIf byCont = 6 Then
                  .ImgMin1(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgMin1(byCont).Visible = True
                  .ImgMin1(byCont).Top = iTop6
               End If
            Else
               .ImgMin1(byCont).Visible = False
               .ImgMin1(byCont).Top = iTopInv
            End If
         
         Next byCont
         
      ElseIf cNombre = "Hor2" Then
      
         For byCont = 0 To 6
            
            If Not Mid$(sSecuencia, byCont + 1, 1) = "0" Then
               If byCont = 0 Then
                  .ImgHora2(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora2(byCont).Visible = True
                  .ImgHora2(byCont).Top = iTop0
               ElseIf byCont = 1 Or byCont = 5 Then
                  .ImgHora2(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora2(byCont).Visible = True
                  .ImgHora2(byCont).Top = iTop15
               ElseIf byCont = 2 Or byCont = 4 Then
                  .ImgHora2(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora2(byCont).Visible = True
                  .ImgHora2(byCont).Top = iTop24
               ElseIf byCont = 3 Then
                  .ImgHora2(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora2(byCont).Visible = True
                  .ImgHora2(byCont).Top = iTop3
               ElseIf byCont = 6 Then
                  .ImgHora2(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora2(byCont).Visible = True
                  .ImgHora2(byCont).Top = iTop6
               End If
            Else
               .ImgHora2(byCont).Visible = False
               .ImgHora2(byCont).Top = iTopInv
            End If
         
         Next byCont
      
      ElseIf cNombre = "Hor1" Then
         
         For byCont = 0 To 6
            
            If Not Mid$(sSecuencia, byCont + 1, 1) = "0" Then
               If byCont = 0 Then
                  .ImgHora1(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora1(byCont).Visible = True
                  .ImgHora1(byCont).Top = iTop0
               ElseIf byCont = 1 Or byCont = 5 Then
                  .ImgHora1(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora1(byCont).Visible = True
                  .ImgHora1(byCont).Top = iTop15
               ElseIf byCont = 2 Or byCont = 4 Then
                  .ImgHora1(byCont).Picture = Frm_Asistencia.ImgUp(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora1(byCont).Visible = True
                  .ImgHora1(byCont).Top = iTop24
               ElseIf byCont = 3 Then
                  .ImgHora1(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora1(byCont).Visible = True
                  .ImgHora1(byCont).Top = iTop3
               ElseIf byCont = 6 Then
                  .ImgHora1(byCont).Picture = Frm_Asistencia.ImgDown(Mid$(sSecuencia, byCont + 1, 1))
                  .ImgHora1(byCont).Visible = True
                  .ImgHora1(byCont).Top = iTop6
               End If
            Else
               .ImgHora1(byCont).Visible = False
               .ImgHora1(byCont).Top = iTopInv
            End If
         
         Next byCont
         
      End If
  
   End With
   
End Sub

Sub RLJAjusta(bOpcion As Boolean)
'Purpose: Resize if pressed CTRL + H
'Author: L124RD K1N6

   If bOpcion Then
      iTop0 = 300   '0
      iTop15 = 420  '120 '8
      iTop24 = 900  '600  '40
      iTop3 = 1320  '1020 '72
      iTop6 = 780   '480 '36
      iTopInv = 200
      iAltForm = 1600
   Else
      iTop0 = 0
      iTop15 = 0
      iTop24 = 32
      iTop3 = 56
      iTop6 = 28
      iTopInv = 200
      iAltForm = 1380
   End If
   
   Frm_Asistencia.Height = iAltForm
   bFirstTime = False
      
End Sub

Sub RLJKeyDown(iKeyCode As Integer, iShiftKey As Integer)

   Dim bShiftDown As Boolean
   Dim bAltDown   As Boolean
   Dim bCtrlDown  As Boolean
   
   Const RECORRE% = 100    'When you press any arrow this is the distance to be moved
   
   Static bFlag As Boolean

   bShiftDown = (iShiftKey And vbShiftMask) > 0
   bAltDown = (iShiftKey And vbAltMask) > 0
   bCtrlDown = (iShiftKey And vbCtrlMask) > 0
   
   If iKeyCode = vbKeyH Then   'END
      If bShiftDown And bCtrlDown And bAltDown And gbDescarga Then
         Unload Frm_Asistencia
         'End
      ElseIf bCtrlDown Then   'RESIZE
         Call RLJAjusta(bFlag)
         bFlag = Not bFlag
      End If
   End If
   
   'You can move the watch with arrows keys and here i validate the form position
   With Frm_Asistencia
      If iKeyCode = vbKeyDown Then
         If Not .Top > Screen.Height - (.Height) Then .Top = .Top + RECORRE
      ElseIf iKeyCode = vbKeyUp Then
         If Not .Top <= 0 Then .Top = .Top - RECORRE
      ElseIf iKeyCode = vbKeyLeft Then
         If Not .Left <= 0 Then .Left = .Left - RECORRE
      ElseIf iKeyCode = vbKeyRight Then
         If Not .Left > Screen.Width - (.Width) Then .Left = .Left + RECORRE
      End If
   End With
End Sub


