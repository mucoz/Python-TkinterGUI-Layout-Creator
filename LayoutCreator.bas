Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Author             : Mustafa Can Öztürk                                                                                  '
'      Purpose            : Converts VB design into Python code                                                                 '
'      Instructions       : 1) Paste the codes inside a user form                                                               '
'                         : 2) On userform, create a botton on top-left corner with the name "button_generate_tkinter"          '
'                         : 3) Background color of the window must be selected from "Palette" section, not "System"             '
'                         : 4) You can put the elements any place you want                                                      '
'                         : 5) The name of the elements must start with the names below (e.g. "label_status")                   '
'                         : 6) For progress bar, you can create a label and name it as "progressbar_..."                        '
'                         : 7) For menu items, you can use label and the name of the label must start with "menu_"              '
'                         : 8) For each sub item for a menu, you need to add "sub_" keyword e.g. "menu_file_sub_new_sub_exit"   '
'                         : 9) To add a separator, add "sep" keyword between menu items e.g. menu_file_sub_new_sep_sub_exit"    '
'      Supported elements : LABEL ("label_")                                                                                    '
'                         : TEXTBOX ("textbox_")                                                                                '
'                         : BUTTON("button_")                                                                                   '
'                         : OPTION("option_")                                                                                   '
'                         : CHECKBOX ("checkbox_")                                                                              '
'                         : COMBOBOX ("combobox_")                                                                              '
'                         : FRAME ("frame_")                                                                                    '
'                         : MULTIPAGE ("multipage_")                                                                            '
'                         : LISTBOX ("listbox_")                                                                                '
'                         : RICHTEXTBOX ("richtextbox_")                                                                        '
'                         : LISTVIEW ("listview_")                                                                              '
'                         : PROGRESSBAR ("progressbar_")                                                                        '
'                         : MENU ("menu_file_sub_new_sep_sub_edit_sub_view_sep_sub_exit") File-> New | Edit View | Exit         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const LEFT_CONSTANT As Double = 1.35
Private Const TOP_CONSTANT As Double = 1.25
Private Const WIDTH_CONSTANT As Double = 1.35
Private Const HEIGHT_CONSTANT As Double = 1.35

Private coll As Collection
Private headers() As Variant
Private element_layouts() As String
Private element_commands() As String

'LABEL          CHECK
'TEXTBOX        CHECK
'BUTTON         CHECK
'OPTION         CHECK
'CHECKBOX       CHECK
'COMBOBOX       CHECK
'FRAME          CHECK
'MULTIPAGE      CHECK
'LISTBOX        CHECK
'RICHTEXTBOX    CHECK
'LISTVIEW       CHECK
'PROGRESSBAR    CHECK
'MENU           CHECK


'On userform, create a botton on top-left corner with the name "button_generate_tkinter"
Private Sub button_generate_tkinter_Click()
    
    Dim i As Long
    Dim c As Control
    Dim parent_name As String
    Dim element_name As String
    Dim current_page As Integer
    Dim optionbox_number As Integer
    Dim menu_number As Integer
    
    'Collect elements
    Set coll = New Collection
    Call collect_elements(Me.Name)
    
    ReDim element_layouts(0 To 0)
    ReDim element_commands(0 To 0)
    
    optionbox_number = 0
    menu_number = 0
    
    Call prepare_headers
    
    If coll.Count = 0 Then
    
        ReDim Preserve headers(0 To UBound(headers) + 1)
        headers(UBound(headers)) = "print(""No element found!"")"
        compile_code 'headers, element_layouts, element_commands
        Exit Sub
        
    End If

    For i = 1 To coll.Count
        
        element_name = Split(coll(i), ":")(0)
        parent_name = Split(coll(i), ":")(1)
        
        If parent_name = Me.Name Then
        
            parent_name = "window"

        End If
        
        element_name = LCase(element_name)
        parent_name = LCase(parent_name)
        
        If Left(element_name, 7) = "textbox" Then
            
            'Create the layout of the element
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 3)
            
            element_layouts(UBound(element_layouts) - 2) = vbNewLine + vbTab + vbTab + "self." + element_name + "_value" + " = tk.StringVar()" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 1) = element_name + " = ttk.Entry(self." + parent_name + ", textvariable=self." + element_name + "_value, width=" + CStr(CInt(Me.Controls(element_name).Width * 0.22)) + ")" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ")" + vbNewLine
            
            'Create the event/command of the element
            ReDim Preserve element_commands(0 To UBound(element_commands) + 2)
            
            element_commands(UBound(element_commands) - 1) = vbNewLine + vbTab + "def " + LCase(element_name) + "_text(self):" + vbNewLine + vbTab + vbTab
            element_commands(UBound(element_commands)) = "return self." + LCase(element_name) + ".get()" + vbNewLine
        
        ElseIf Left(LCase(element_name), 6) = "button" Then
        
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
    
            element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + element_name + " = ttk.Button(self." + parent_name + ", text='" + Me.Controls(element_name).Caption + "', command=self." + element_name + "_onclick)" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine
            
            ReDim Preserve element_commands(0 To UBound(element_commands) + 2)
            
            element_commands(UBound(element_commands) - 1) = vbNewLine + vbTab + "def " + element_name + "_onclick(self):" + vbNewLine + vbTab + vbTab
            element_commands(UBound(element_commands)) = "print('" + element_name + " has been clicked')" + vbNewLine
        
        ElseIf Left(element_name, 5) = "label" Then
        
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
            
            element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + element_name + " = ttk.Label(self." + parent_name + ", text='" + Me.Controls(element_name).Caption + "', background=self.background_color)" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ")" + vbNewLine
        
        ElseIf Left(element_name, 8) = "checkbox" Then
        
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 4)
            
            element_layouts(UBound(element_layouts) - 3) = vbNewLine + vbTab + vbTab + "self." + element_name + "_value = tk.IntVar()" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 2) = "self." + element_name + " = tk.Checkbutton(self." + parent_name + ", text='" + Me.Controls(element_name).Caption + "', variable=self." + element_name + "_value, background=self.background_color)" + vbNewLine + vbTab + vbTab
            If Me.Controls(element_name).value = True Then
                element_layouts(UBound(element_layouts) - 1) = "self." + element_name + ".select()" + vbNewLine + vbTab + vbTab
            Else
                element_layouts(UBound(element_layouts) - 1) = "self." + element_name + ".deselect()" + vbNewLine + vbTab + vbTab
            End If
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ")" + vbNewLine
            
        ElseIf Left(element_name, 5) = "frame" Then
            
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
            
            element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + element_name + " = ttk.LabelFrame(self." + parent_name + ", text='" + Me.Controls(element_name).Caption + "')" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine
            optionbox_number = 0
            
        ElseIf Left(element_name, 8) = "combobox" Then
        
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 4)
            
            element_layouts(UBound(element_layouts) - 3) = vbNewLine + vbTab + vbTab + "self." + element_name + "_value = tk.StringVar()" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 2) = "self." + element_name + " = ttk.Combobox(self." + parent_name + ", textvariable=self." + element_name + "_value, state='readonly')" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 1) = "self." + element_name + "['values'] = ('Item1', 'Item2', 'Item3')"
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine
            
        ElseIf Left(element_name, 11) = "richtextbox" Then
            
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
            
            element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + element_name + " = scrolledtext.ScrolledText(self." + parent_name + ", wrap=tk.WORD)" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine
            
            ReDim Preserve element_commands(0 To UBound(element_commands) + 2)
            
            element_commands(UBound(element_commands) - 1) = vbNewLine + vbTab + "def " + element_name + "_text(self):" + vbNewLine + vbTab + vbTab
            element_commands(UBound(element_commands)) = "return self." + element_name + ".get('1.0', tk.END)" + vbNewLine
            
        ElseIf Left(element_name, 9) = "multipage" Then
        
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
            
            element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + element_name + " = ttk.Notebook(self." + parent_name + ")" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine
            current_page = 0
            optionbox_number = 0
        
        ElseIf Left(element_name, 4) = "page" Then
        
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
            
            element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + element_name + " = ttk.Frame(self." + parent_name + ")" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = "self." + parent_name + ".add(self." + element_name + ", text='" + Me.Controls(parent_name).Pages(current_page).Caption + "')" + vbNewLine
            current_page = current_page + 1
            
        ElseIf Left(element_name, 7) = "listbox" Then
            
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 15)
            
            element_layouts(UBound(element_layouts) - 14) = vbNewLine + "#################### " + UCase(element_name) + ", (Including Horizontal and Vertical Scrollbars) ####################" + vbNewLine + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 13) = "self." + element_name + "_values = ('Item1', 'Item2', 'Item3')" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 12) = "self.values_" + element_name + " = tk.Variable(value=self." + element_name + "_values)" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 11) = "self." + element_name + " = tk.Listbox(self." + parent_name + ", listvariable=self.values_" + element_name + ")" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 10) = "self." + element_name + ".bind('<Double-1>', self." + element_name + "_doubleclick) # you can disable this functionality" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 9) = "self." + element_name + ".bind('<<ListboxSelect>>', self." + element_name + "_onclick) # you can disable this functionality" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 8) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 7) = "self." + element_name + "_scrollbarx = ttk.Scrollbar(self." + parent_name + ", orient=tk.HORIZONTAL, command=self." + element_name + ".xview())" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 6) = "self." + element_name + "_scrollbarx.place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT) + CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=20)" + vbNewLine + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 5) = "self." + element_name + "_scrollbary = ttk.Scrollbar(self." + parent_name + ", orient=tk.VERTICAL, command=self." + element_name + ".yview())" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 4) = "self." + element_name + "_scrollbary.place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT) + CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=20, height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 3) = "self." + element_name + ".config(xscrollcommand=self." + element_name + "_scrollbarx.set, yscrollcommand=self." + element_name + "_scrollbary.set)" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 2) = "self." + element_name + "_scrollbarx.config(command=self." + element_name + ".xview)" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 1) = "self." + element_name + "_scrollbary.config(command=self." + element_name + ".yview)" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = "#################### END OF " + UCase(element_name) + ", (Including Horizontal and Vertical Scrollbars) ####################" + vbNewLine
            
            ReDim Preserve element_commands(0 To UBound(element_commands) + 6)
            
            element_commands(UBound(element_commands) - 5) = vbNewLine + vbTab + "def " + element_name + "_doubleclick(self, event):" + vbNewLine + vbTab + vbTab
            element_commands(UBound(element_commands) - 4) = "cs = self." + element_name + ".curselection()" + vbNewLine + vbTab + vbTab
            element_commands(UBound(element_commands) - 3) = "print('Double clicked on ' + self." + element_name + ".get(cs))" + vbNewLine + vbNewLine + vbTab
            element_commands(UBound(element_commands) - 2) = "def " + element_name + "_onclick(self, event):" + vbNewLine + vbTab + vbTab
            element_commands(UBound(element_commands) - 1) = "cs = self." + element_name + ".curselection()" + vbNewLine + vbTab + vbTab
            element_commands(UBound(element_commands)) = "print('Clicked on ' + self." + element_name + ".get(cs))" + vbNewLine
            
        
        ElseIf Left(element_name, 8) = "listview" Then
        
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 10)
            
            element_layouts(UBound(element_layouts) - 9) = vbNewLine + vbTab + vbTab + "self." + element_name + "_column_names = ('Column 1', 'Column 2', 'Column 3')" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 8) = "self." + element_name + " = ttk.Treeview(self." + parent_name + ", columns=self." + element_name + "_column_names, show='headings')" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 7) = "for i in range(len(self." + element_name + "['columns'])):" + vbNewLine + vbTab + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 6) = "self." + element_name + ".column(self." + element_name + "['column'][i], anchor=tk.CENTER, width=150)" + vbNewLine + vbTab + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 5) = "self." + element_name + ".heading(self." + element_name + "['columns'][i], text=self." + element_name + "['columns'][i], anchor=tk.CENTER)" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 4) = "self." + element_name + ".bind('<Double-1>', self." + element_name + "_doubleclick)" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts) - 3) = "self." + element_name + ".insert(parent='', index=i, text='', values=([1, 2, 3]))"
            element_layouts(UBound(element_layouts) - 2) = "self." + element_name + ".insert(parent='', index=i, text='', values=([4, 5, 6]))"
            element_layouts(UBound(element_layouts) - 1) = "self." + element_name + ".insert(parent='', index=i, text='', values=([7, 8, 9]))"
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine
        
        
            'Create double click event
            ReDim Preserve element_commands(0 To UBound(element_commands) + 4)
            
            element_commands(UBound(element_commands) - 3) = vbNewLine + vbTab + "def " + element_name + "_doubleclick(self, event):" + vbNewLine + vbTab + vbTab
            element_commands(UBound(element_commands) - 2) = "item = self." + element_name + ".selection()" + vbNewLine + vbTab + vbTab
            element_commands(UBound(element_commands) - 1) = "for i in item:" + vbNewLine + vbTab + vbTab + vbTab
            element_commands(UBound(element_commands)) = "print('Double clicked on '," + "self." + element_name + ".item(i, 'values')[0])" + vbNewLine
            
        ElseIf Left(element_name, 6) = "option" Then
            
            If optionbox_number = 0 Then
                
                ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
                
                element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self.options_" + parent_name + "_value = tk.IntVar()" + vbNewLine + vbTab + vbTab
                element_layouts(UBound(element_layouts)) = "self.options_" + parent_name + "_value.set(99)" + vbNewLine
                
            End If
            
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
            
            element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + element_name + " = tk.Radiobutton(self." + parent_name + ", text='" + Me.Controls(element_name).Caption + "', bg=self.background_color, variable=self.options_" + parent_name + "_value, value=" + CStr(optionbox_number) + ")" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine
             
            optionbox_number = optionbox_number + 1
            
        ElseIf Left(element_name, 11) = "progressbar" Then
        
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
            
            element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self." + element_name + " = ttk.Progressbar(self." + parent_name + ", orient='horizontal', length=" + CStr(Me.Controls(element_name).Width) + ", mode='determinate')" + vbNewLine + vbTab + vbTab
            element_layouts(UBound(element_layouts)) = "self." + element_name + ".place(x=" + CStr(CInt(Me.Controls(element_name).Left * LEFT_CONSTANT)) + ", y=" + CStr(CInt(Me.Controls(element_name).Top * TOP_CONSTANT)) + ", width=" + CStr(CInt(Me.Controls(element_name).Width * WIDTH_CONSTANT)) + ", height=" + CStr(CInt(Me.Controls(element_name).Height * HEIGHT_CONSTANT)) + ")" + vbNewLine
        
        ElseIf Left(element_name, 4) = "menu" Then
            
            Dim j As Integer
            Dim menu_coll As Collection
            Dim a As Variant
            
            Set menu_coll = New Collection
            
            'Create this block of code once for menu creation
            If menu_number = 0 Then
            
                ReDim Preserve element_layouts(0 To UBound(element_layouts) + 2)
                
                element_layouts(UBound(element_layouts) - 1) = vbNewLine + vbTab + vbTab + "self.menu_bar = Menu(self.window)" + vbNewLine + vbTab + vbTab
                element_layouts(UBound(element_layouts)) = "self.window.config(menu=self.menu_bar)" + vbNewLine
            
                menu_number = menu_number + 1
                
            End If
            
            a = Split(element_name, "_")
            
            For j = LBound(a) To UBound(a)
                
                If LCase(a(j)) <> "menu" And LCase(a(j)) <> "sub" Then
                    
                    menu_coll.Add cUpper(a(j)), CStr(j)
                         
                End If
            
            Next j
            
            For j = 1 To menu_coll.Count
            
                If j = 1 Then
                
                    ReDim Preserve element_layouts(0 To UBound(element_layouts) + 1)
                    element_layouts(UBound(element_layouts)) = vbNewLine + vbTab + vbTab + "self." + LCase(menu_coll(j)) + "_menu = Menu(self.menu_bar, tearoff=0)"
                
                Else
                    
                    If LCase(menu_coll(j)) = "sep" Then
                                            
                        ReDim Preserve element_layouts(0 To UBound(element_layouts) + 1)
                        element_layouts(UBound(element_layouts)) = vbNewLine + vbTab + vbTab + "self." + LCase(menu_coll(1)) + "_menu.add_separator()"
                    
                    Else
                    
                        ReDim Preserve element_layouts(0 To UBound(element_layouts) + 1)
                        element_layouts(UBound(element_layouts)) = vbNewLine + vbTab + vbTab + "self." + LCase(menu_coll(1)) + "_menu.add_command(label='" + menu_coll(j) + "', command=self.menu_" + LCase(menu_coll(j)) + "_click)"
                    
                        ReDim Preserve element_commands(0 To UBound(element_commands) + 2)
                
                        element_commands(UBound(element_commands) - 1) = vbNewLine + vbTab + "def menu_" + LCase(menu_coll(j)) + "_click(self):" + vbNewLine + vbTab + vbTab
                        element_commands(UBound(element_commands)) = "print('menu_" + LCase(menu_coll(j)) + " has been clicked')" + vbNewLine
                                    
                    End If
                    
                End If
            
            Next j
            
            ReDim Preserve element_layouts(0 To UBound(element_layouts) + 1)
            
            element_layouts(UBound(element_layouts)) = vbNewLine + vbTab + vbTab + "self.menu_bar.add_cascade(label='" + menu_coll(1) + "', menu=self." + LCase(menu_coll(1)) + "_menu)" + vbNewLine
            
        End If
    
    Next i
    
    Call compile_code
    
End Sub

Private Sub prepare_headers()

    ' At the beginning, create the default layout
    ReDim headers(0 To 15)
    
    headers(0) = "import tkinter as tk" + vbNewLine
    headers(1) = "from tkinter import ttk" + vbNewLine
    headers(2) = "from tkinter import scrolledtext" + vbNewLine
    headers(3) = "from tkinter import Menu" + vbNewLine
    headers(4) = "from tkinter import messagebox as msg" + vbNewLine
    headers(5) = vbNewLine
    headers(6) = "class new_window():" + vbNewLine + vbTab
    headers(7) = "def __init__(self):" + vbNewLine + vbTab + vbTab
    headers(8) = "self.window = tk.Tk()" + vbNewLine + vbTab + vbTab
    headers(9) = "self.window.title ('" + Me.Caption + "')" + vbNewLine + vbTab + vbTab
    headers(10) = "self.window.geometry ('" + CStr(CInt(Me.Width * 1.315)) + "x" + CStr(CInt(Me.Height * 1.265)) + "')" + vbNewLine + vbTab + vbTab
    headers(11) = "self.background_color='#" + get_bgcolor_hex + "'" + vbNewLine + vbTab + vbTab
    headers(12) = "self.window.configure(bg=self.background_color)" + vbNewLine + vbTab + vbTab
    headers(13) = "self.create_elements()" + vbNewLine + vbTab + vbTab
    headers(14) = "self.window.mainloop()" + vbNewLine + vbNewLine
    headers(15) = vbTab + "def create_elements(self):" + vbNewLine + vbTab + vbTab

End Sub

Private Sub compile_code()
    
    Dim i As Long
    Dim text As String
    
    For i = LBound(headers) To UBound(headers)
        
        text = text + headers(i)

    Next i
    
    If coll.Count <> 0 Then
        For i = LBound(element_layouts) To UBound(element_layouts)
            
            text = text + element_layouts(i)
            
        Next i
        
        For i = LBound(element_commands) To UBound(element_commands)
        
            text = text + element_commands(i)
        
        Next i
    End If
    
    SetClipboard text
    
End Sub

Private Sub collect_elements(element_name)

    Dim c As Control
    Dim i As Long
    
    For Each c In Me.Controls
        If c.Name <> "button_generate_tkinter" Then
            If c.Parent.Name = element_name Then
                coll.Add c.Name + ":" + c.Parent.Name, c.Name
                If LCase(TypeName(c)) = "multipage" Then
                    For i = 0 To c.Pages.Count - 1
                        coll.Add c.Pages(i).Name + ":" + c.Pages(i).Parent.Name, c.Pages(i).Name
                        Call collect_elements(c.Pages(i).Name)
                    Next i
                Else
                    Call collect_elements(c.Name)
                End If
            End If
        End If
    Next c
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

Private Function cUpper(ByVal value As String) As String

    cUpper = UCase(Mid(value, 1, 1)) + Mid(value, 2, Len(value) - 1)
        
End Function

Private Sub SetClipboard(text As String)

    Dim obj As New DataObject
    obj.SetText text
    obj.PutInClipboard

End Sub

