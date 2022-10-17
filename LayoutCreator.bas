Option Explicit

'On userform, create a botton on top-left corner with the name "button_generate_tkinter"
Private Sub button_generate_tkinter_Click()
    
    Dim str() As String
    Dim text As String
    Dim i As Long
    Dim c As Control
    
    ' At the beginning, create the default layout
    ReDim str(0 To 7)
    
    str(0) = "import tkinter as tk" + vbNewLine
    str(1) = "from tkinter import ttk" + vbNewLine
    str(2) = vbNewLine
    str(3) = "win = tk.Tk()" + vbNewLine
    str(4) = "win.title ('" + Me.Caption + "')" + vbNewLine
    str(5) = "win.geometry ('" + CStr(CInt(Me.Width * 1.315)) + "x" + CStr(CInt(Me.Height * 1.265)) + "')" + vbNewLine
    str(6) = "win.configure(bg='#" + get_bgcolor_hex + "')" + vbNewLine + vbNewLine
    str(7) = "win.mainloop()" + vbNewLine
    
    'start adding the elements on the userform if there is any
    
    For Each c In Me.Controls
        
        If c.Name <> "button_generate_tkinter" Then
            
            If InStr(LCase(c.Name), "textbox") > 0 Then
                
                ReDim Preserve str(0 To UBound(str) + 5)
                str(UBound(str) - 5) = LCase(c.Name) + " = tk.StringVar()" + vbNewLine
                str(UBound(str) - 4) = LCase(c.Name) + "_control = ttk.Entry(win, textvariable=" + LCase(c.Name) + ", width=" + CStr(CInt(c.Width * 0.22)) + ")" + vbNewLine
                str(UBound(str) - 3) = LCase(c.Name) + "_control.place(x=" + CStr(CInt(c.Left * 1.25)) + ", y=" + CStr(c.Top) + ")" + vbNewLine
                str(UBound(str) - 2) = "def " + LCase(c.Name) + "_text():" + vbNewLine
                str(UBound(str) - 1) = vbTab + "return " + LCase(c.Name) + ".get()" + vbNewLine + vbNewLine
                str(UBound(str)) = "win.mainloop()" + vbNewLine
            
            ElseIf InStr(LCase(c.Name), "button") > 0 Then
            
                ReDim Preserve str(0 To UBound(str) + 5)
                                
                str(UBound(str) - 5) = "def " + LCase(c.Name) + "_onclick():" + vbNewLine
                str(UBound(str) - 4) = vbTab + "print('" + LCase(c.Name) + " has been clicked')" + vbNewLine
                str(UBound(str) - 3) = vbNewLine
                str(UBound(str) - 2) = LCase(c.Name) + " = ttk.Button(win, text='" + c.Caption + "', command=" + LCase(c.Name) + "_onclick)" + vbNewLine
                str(UBound(str) - 1) = LCase(c.Name) + ".place(x=" + CStr(CInt(c.Left * 1.25)) + ", y=" + CStr(c.Top) + ", width=" + CStr(CInt(c.Width * 1.5)) + ", height=" + CStr(CInt(c.Height * 1.5)) + ")" + vbNewLine + vbNewLine
                str(UBound(str)) = "win.mainloop()" + vbNewLine
            
            ElseIf InStr(LCase(c.Name), "label") > 0 Then
            
                ReDim Preserve str(0 To UBound(str) + 2)
                
                str(UBound(str) - 2) = LCase(c.Name) + " = ttk.Label(win, text='" + c.Caption + "', background='#" + get_bgcolor_hex + "')" + vbNewLine
                str(UBound(str) - 1) = LCase(c.Name) + ".place(x=" + CStr(CInt(c.Left)) + ", y=" + CStr(c.Top) + ")" + vbNewLine + vbNewLine
                str(UBound(str)) = "win.mainloop()" + vbNewLine
            
            End If
        
        End If
    
    Next c
    
    For i = LBound(str) To UBound(str)
    
        text = text + str(i)
    
    Next i
    
    
    SetClipboard text
    
End Sub

Private Function get_bgcolor_hex() As String

    Dim FillHexColor As String
    Dim r As String, g As String, b As String

    'Get Hex values (values come through in reverse of what we need)
    FillHexColor = Right("000000" & Hex(Me.BackColor), 6)
        If Len(FillHexColor) > 4 Then
            r = Right(FillHexColor, 2)
            g = Mid(FillHexColor, 3, 2)
            b = Left(FillHexColor, 2)
        Else
            r = r = Right(FillHexColor, 2)
            g = Left(FillHexColor, 2)
            b = "00"
        End If
        
        FillHexColor = r + g + b
    
    get_bgcolor_hex = FillHexColor

End Function

Private Sub SetClipboard(text As String)

    Dim obj As New DataObject
    obj.SetText text
    obj.PutInClipboard

End Sub
