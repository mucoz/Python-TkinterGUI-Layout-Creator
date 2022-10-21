Option Explicit

Private Const LEFT_CONSTANT As Double = 1.25
Private Const TOP_CONSTANT As Double = 1.25
Private Const WIDTH_CONSTANT As Double = 1.35
Private Const HEIGHT_CONSTANT As Double = 1.35

'LABEL          CHECK
'TEXTBOX        CHECK
'BUTTON         CHECK
'OPTIONBOX
'CHECKBOX       CHECK
'COMBOBOX
'FRAME
'MULTIPAGE
'LISTBOX
'RICHTEXTBOX
'LISTVIEW
'MENU


'On userform, create a botton on top-left corner with the name "button_generate_tkinter"
Private Sub button_generate_tkinter_Click()
    
    Dim str() As String
    Dim element_layouts() As String
    Dim element_commands() As String
    Dim text As String
    Dim i As Long
    Dim c As Control
    Dim counter As Long
    Dim p As Page
    Dim pageName As String
    Dim pageCaption As String
    Dim parentWindow As String
    
    ReDim element_layouts(0 To 0)
    ReDim element_commands(0 To 0)
    ' At the beginning, create the default layout
    ReDim str(0 To 13)
    
    str(0) = "import tkinter as tk" + vbNewLine
    str(1) = "from tkinter import ttk" + vbNewLine
    str(2) = "from tkinter import Menu" + vbNewLine
    str(3) = "from tkinter import messagebox as msg" + vbNewLine
    str(4) = vbNewLine
    str(5) = "class new_window():" + vbNewLine + vbTab
    str(6) = "def __init__(self):" + vbNewLine + vbTab + vbTab
    str(7) = "self.window = tk.Tk()" + vbNewLine + vbTab + vbTab
    str(8) = "self.window.title ('" + Me.Caption + "')" + vbNewLine + vbTab + vbTab
    str(9) = "self.window.geometry ('" + CStr(CInt(Me.Width * 1.315)) + "x" + CStr(CInt(Me.Height * 1.265)) + "')" + vbNewLine + vbTab + vbTab
    str(10) = "self.background_color='#" + get_bgcolor_hex + "'" + vbNewLine + vbTab + vbTab
    str(11) = "self.window.configure(bg=self.background_color)" + vbNewLine + vbTab + vbTab
    str(12) = "self.create_elements()" + vbNewLine + vbTab + vbTab
    str(13) = "self.window.mainloop()" + vbNewLine + vbNewLine
    
    'start adding the elements on the userform if there is any
    ReDim Preserve str(UBound(str) + 1)
    str(UBound(str)) = vbTab + "def create_elements(self):" + vbNewLine + vbTab + vbTab

    
    
    'Check if we have tab control
    For Each c In Me.Controls
            
        If LCase(TypeName(c)) = "multipage" Then
            
            ReDim Preserve element_layouts(UBound(element_layouts) + 1)
            
            If Left(LCase(c.Parent.Name), 6) = "window" Then
                parentWindow = "window"
            ElseIf Left(LCase(c.Parent.Name), 4) = "page" Then
                parentWindow = "page"
            ElseIf Left(LCase(c.Parent.Name), 5) = "frame" Then
                parentWindow = "frame"
            End If
            
            element_layouts(UBound(element_layouts)) = LCase(c.Name) + " = ttk.Notebook(self." + parentWindow + ")" + vbNewLine + vbTab + vbTab
            
            For i = 0 To c.Pages.Count - 1
                pageName = c.Pages(i).Name
                pageCaption = c.Pages(i).Caption
                
                ReDim Preserve element_layouts(UBound(element_layouts) + 2)
                
                element_layouts(UBound(element_layouts) - 1) = pageName + " = ttk.Frame(" + LCase(c.Name) + ")" + vbNewLine + vbTab + vbTab
                element_layouts(UBound(element_layouts)) = LCase(c.Name) + ".add(" + pageName + ", text='" + pageCaption + "')" + vbNewLine + vbTab + vbTab
                
            Next i
            
            ReDim Preserve element_layouts(UBound(element_layouts) + 1)
            
            element_layouts(UBound(element_layouts)) = LCase(c.Name) + ".place(x=" + CStr(CInt(c.Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(c.Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(c.Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(c.Height * HEIGHT_CONSTANT)) + ")" + vbNewLine + vbTab + vbTab

        End If
        
    Next c
    
    
    'Check if we have frames
    For Each c In Me.Controls
    
    
    
    Next c
    
    'Check other control objects
    For Each c In Me.Controls
        
        If c.Name <> "button_generate_tkinter" Then
            
            'Count the number of elements on the form
            counter = counter + 1
            
            If Left(LCase(c.Name), 7) = "textbox" Then
                
                'Create the layout of the element
                ReDim Preserve element_layouts(0 To UBound(element_layouts) + 3)
                
                element_layouts(UBound(element_layouts) - 2) = vbNewLine + vbTab + vbTab + "self." + LCase(c.Name) + " = tk.StringVar()" + vbNewLine + vbTab + vbTab
                element_layouts(UBound(element_layouts) - 1) = LCase(c.Name) + "_control = ttk.Entry(self.window, textvariable=self." + LCase(c.Name) + ", width=" + CStr(CInt(c.Width * 0.22)) + ")" + vbNewLine + vbTab + vbTab
                element_layouts(UBound(element_layouts)) = LCase(c.Name) + "_control.place(x=" + CStr(CInt(c.Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(c.Top * TOP_CONSTANT)) + ")" + vbNewLine
                
                'Create the event/command of the element
                ReDim Preserve element_commands(0 To UBound(element_commands) + 2)
                
                element_commands(UBound(element_commands) - 1) = vbNewLine + vbTab + "def " + LCase(c.Name) + "_text(self):" + vbNewLine + vbTab + vbTab
                element_commands(UBound(element_commands)) = "return self." + LCase(c.Name) + ".get()" + vbNewLine
            
            ElseIf Left(LCase(c.Name), 6) = "button" Then
            
                ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
                                
  
                element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + LCase(c.Name) + " = ttk.Button(self.window, text='" + c.Caption + "', command=self." + LCase(c.Name) + "_onclick)" + vbNewLine + vbTab + vbTab
                element_layouts(UBound(element_layouts)) = "self." + LCase(c.Name) + ".place(x=" + CStr(CInt(c.Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(c.Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(c.Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(c.Height * HEIGHT_CONSTANT)) + ")" + vbNewLine
                
                ReDim Preserve element_commands(0 To UBound(element_commands) + 2)
                
                element_commands(UBound(element_commands) - 1) = vbNewLine + vbTab + "def " + LCase(c.Name) + "_onclick(self):" + vbNewLine + vbTab + vbTab
                element_commands(UBound(element_commands)) = "print('" + LCase(c.Name) + " has been clicked')" + vbNewLine
            
            ElseIf Left(LCase(c.Name), 5) = "label" Then
            
                ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
                
                element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + LCase(c.Name) + " = ttk.Label(self.window, text='" + c.Caption + "', background=self.background_color)" + vbNewLine + vbTab + vbTab
                element_layouts(UBound(element_layouts)) = "self." + LCase(c.Name) + ".place(x=" + CStr(CInt(c.Left)) + ", y=" + CStr(CInt(c.Top * TOP_CONSTANT)) + ")" + vbNewLine
            
            ElseIf Left(LCase(c.Name), 8) = "checkbox" Then
            
                ReDim Preserve element_layouts(0 To UBound(element_layouts) + 4)
                
                element_layouts(UBound(element_layouts) - 3) = vbNewLine + vbTab + vbTab + "self." + LCase(c.Name) + "_value = tk.IntVar()" + vbNewLine + vbTab + vbTab
                element_layouts(UBound(element_layouts) - 2) = "self." + LCase(c.Name) + " = tk.Checkbutton(self.window, text='" + c.Caption + "', variable=self." + LCase(c.Name) + "_value, background=self.background_color)" + vbNewLine + vbTab + vbTab
                If c.Value = True Then
                    element_layouts(UBound(element_layouts) - 1) = "self." + LCase(c.Name) + ".select()" + vbNewLine + vbTab + vbTab
                Else
                    element_layouts(UBound(element_layouts) - 1) = "self." + LCase(c.Name) + ".deselect()" + vbNewLine + vbTab + vbTab
                End If
                element_layouts(UBound(element_layouts)) = "self." + LCase(c.Name) + ".place(x=" + CStr(CInt(c.Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(c.Top * TOP_CONSTANT)) + ")" + vbNewLine
                
            ElseIf Left(LCase(c.Name), 6) = "option" Then
                
                ReDim Preserve element_layouts(0 To UBound(element_layouts) + 3)
                
            
            
            End If
        
        End If
    
    Next c
    
    If counter = 0 Then
    
        ReDim Preserve str(0 To UBound(str) + 1)
        
        'str(UBound(str) - 1) = vbTab + "def create_elements(self):" + vbNewLine + vbTab + vbTab
        str(UBound(str)) = "print(""No element found!"")"
        
    End If
    
    For i = LBound(str) To UBound(str)
        
        text = text + str(i)

    Next i
    
    For i = LBound(element_layouts) To UBound(element_layouts)
        
        text = text + element_layouts(i)
        
    Next i
    
    For i = LBound(element_commands) To UBound(element_commands)
    
        text = text + element_commands(i)
    
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


Private Sub button2_Click()

End Sub
