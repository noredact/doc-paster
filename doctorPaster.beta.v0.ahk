﻿

#Requires AutoHotkey v2+
#SingleInstance
#Warn
SendMode "Input"
SetWorkingDir A_ScriptDir
TraySetIcon A_ScriptDir . "\media\icon.ico",, true
; Read initialization settings from INI
filePath := IniRead(A_ScriptFullPath . ":Stream:$DATA", "initialization", "defaultFilePath", A_ScriptDir . "\Lorem.csv")
dirPath := IniRead(A_ScriptFullPath . ":Stream:$DATA", "initialization", "defaultDirPath", A_ScriptDir)
defaultFontSize := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "fontSize", "8")
defaultEmptyString := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "defaultEmptyString", "¯\_(ツ)_/¯")
loadLastWinPos := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "loadLastWinPos", "1")
loadLastFile := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "loadLastFile", "0")
loadLastPos := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "loadLastPosition", "0")
lastPosition := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "lastPosition", "0")
loadViewSettings := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "loadViewSettings", "0")
textScroll := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "textScroll", "0")
windowTransparency := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "windowTransparency", "255")
textTransparency := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "textTransparency", "0")
lastWinPos := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "lastWinPos", "x0 y0")
showSettings := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "showSettings", "0")

; Global function for reading target file
readTargetFile(filePath) {
; msgbox filePath
    contentObj := {}
    contentArray := []
    lengthArray := []
    lineArray := []
    emptyElementStr := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "defaultEmptyString", "¯\_(ツ)_/¯")
    contentsRaw := FileRead(filePath)
    fileName := filePath
    endofLine := 0
    endofLineTotal := 0

    contentsRaw := RegExReplace(contentsRaw, "\R", "`n")

    ; Split into lines and filter out empty lines
    lines := StrSplit(contentsRaw, "`n")
    filteredLines := []

    ; Filter out completely empty lines
    for _, line in lines {
        if (Trim(line) != "") {
            filteredLines.Push(line)
        }
    }

    for _, currentLine in filteredLines {
        Loop Parse currentLine, "CSV" {
            endofLine := A_Index
            readElement := A_LoopField
            if readElement == ""
                readElement := emptyElementStr
            lengthArray.Push(StrLen(readElement))
            contentArray.Push(readElement)
        }
        endofLineTotal += endofLine
        lineArray.Push(endofLineTotal)
    }

    checkContentLength := contentArray.Length

    ; Ensure 9 elements, Content array can't have length less than 9
    if checkContentLength < 9 {
        fillCount := 9 - checkContentLength
        Loop fillCount {
            contentArray.push(emptyElementStr)
        }
    }

    elementCount := contentArray.Length
    contentObj := {
        contentArr: contentArray,
        lengthArr: lengthArray,
        eleCount: elementCount,
        lines: filteredLines.Length,
        lineArr: lineArray,
        name: fileName,
    }
    return contentObj
}


; Main GUI class
class drPaster {
    currentToggle := showSettings == 1 ? showSettings : false
    currentPosition := 0
    currentLine := 1
    incAmount := 1
    previewSize := 3
    dpasteGui := ""
    curFile := ""
    scrollArray := []
    textLengthThreshold := 12  ; Maximum length before scrolling
    optionTextLengthThreshold := 12
    scrollTimer := ""

    __New() {
        this.dpasteGui := Gui()
        this.filePath := loadLastFile == 1 ? filePath : FileSelect()
        this.curFile := readTargetFile(filePath)
        this.fileName := this.curFile.name
        this.lastPos := loadLastPos == 1 ? lastPosition : 0
        this.winTransparency := loadViewSettings == 1 ? windowTransparency : 255
        this.textTransparency := loadViewSettings == 1 ? textTransparency : 0
        this.contentLength := this.curFile.contentArr.Length
        this.lineCounterArr := this.curFile.lineArr
        this.totalLines := this.curFile.lines
        this.parentDir := ""
        this.grandParentDir := ""
        this.TVdir := ""
        if RegExMatch(this.fileName, "[\\/](?<GrandParent>[^\\/]+)[\\/][^\\/]+[\\/][^\\/]+$", &match) {
            this.grandParentDir := match[1]
        }
        if RegExMatch(this.fileName, "[\\/](?<Parent>[^\\/]+)[\\/][^\\/]+$", &match) {
            this.parentDir := match[1]
        }
        ; RegExReplace(this.fileName, "^(.*)[\\/][^\\/]+$", this.parentDir)
        ; ; this.parentDir := RegExReplace(this.fileName, "^(.*)[\\/][^\\/]+[\\/]?$", "")
        this.TVdir := "\" this.grandParentDir . "\" . this.parentDir . "\"
        
        ; msgbox this.TVdir
    
        ; Prevent GUI from opening out of bounds
        if lastWinPos != "" 
        {
            VirtualScreenWidth := SysGet(78)
            VirtualScreenHeight := SysGet(79)
            displaySize := [VirtualScreenWidth, VirtualScreenHeight]
            posOutOfBoundsCheck := StrSplit(lastWinPos, " ", "-xy")
            if (displaySize[1] < posOutOfBoundsCheck[1] || displaySize[2] < posOutOfBoundsCheck[2])
                this.winPos := ""
            else
                this.winPos := lastWinPos
        }
            
        ; Menu Setup
        FileMenu := Menu()
        FileMenu.Add("&Open Directory`tCtrl+o", (*) => this.changeDir())
        FileMenu.Add("E&xit", (*) => this.exitApp())

        SettingsMenu := Menu()
        SettingsMenu.Add("Show/hide Optio&ns`tCtrl+n", (*) => this.settingsToggle())
        SettingsMenu.Add("&Font Size`tCtrl+F", (*) => this.showFontSizeDialog())
        SettingsMenu.Add("&Jump to...`tCtrl+J", (*) => this.showJumpPosDialog())
        SettingsMenu.Add("Change default empty field string", (*) => this.showEmptyStringDialog())

        HelpMenu := Menu()
        HelpMenu.Add("&About`tF1", (*) => this.showAbout())

        Menus := MenuBar()
        Menus.Add("&File", FileMenu)
        Menus.Add("&Settings", SettingsMenu)
        Menus.Add("&Help", HelpMenu)
        this.dpasteGui.MenuBar := Menus

        this.initializeDisplay()
        this.addSettingsSection()
        this.addPreviewSection()
        this.addStatusBar()
        this.addDirectoryView()
        this.addUserOptions()

        
        ; Initialize scroll timer
        this.scrollTimer := ObjBindMethod(this, "moveText")
        if (this.dpasteGui["textScroll"].Value == 1) {
            SetTimer(this.scrollTimer, 240)
        } else {
            SetTimer(this.scrollTimer, 0)
        }

        this.loadSettings()

        ; this.dpasteGui.Show("w1183" this.winPos)
        this.dpasteGui.Show(this.winPos)
        this.dpasteGui.Opt("+AlwaysOnTop")
        this.settingsToggle()
        this.guiHwnd := this.dpasteGui.Hwnd
        this.dpasteGui.OnEvent('Close', (*) => this.exitApp())
        this.dpasteGui.OnEvent('Escape', (*) => this.exitApp())

        OnExit(ObjBindMethod(this, "handleExit"))
    }




    initializeDisplay() {
        Loop 9 {
            currentArrayPos := this.currentPosition + A_Index
            if (A_Index = 1) {
                this.dpasteGui.Add("GroupBox", "Center Section xm y3 w125 h55", A_Index)
            } else {
                this.dpasteGui.Add("GroupBox", "Center Section xs+130 y3 w125 h55", A_Index)
            }
            
            textContent := this.curFile.contentArr[currentArrayPos]
            this.dpasteGui.Add("Text", "xs+5 ys+23 w115 h15 -Wrap Center  vText" A_Index, textContent)
            
            ; Add to scroll array if text is too long
            if (StrLen(textContent) > this.textLengthThreshold) {
                this.scrollArray.Push("Text" A_Index)
            }
        }

    }
	
    addSettingsSection() {
        this.dpasteGui.Add("GroupBox", "Section xm+2 y+24 h140 w275", "Display and navigation settings")




        ; Transparency Controls
        this.dpasteGui.Add("GroupBox", "Section xp+5 yp+15 hp-20",  "Visiblity ctrl+shift+Num(+/-)")
        this.dpasteGui.Add("Text","xp+1 ys+15 Center wp-2","<< Less --------------- More >>")
        this.dpasteGui.Add("Text", "Center xp+5 yp+18", "Window ctrl+Num(+/-)")
        this.dpasteGui.Add("Slider", "vWindowTransparency NoTicks Center AltSubmit ToolTipTop Range50-255 xp yp+20 h20", "255")

        this.dpasteGui.Add("Text", "Center xp yp+30", "Text shift+Num(+/-)")
        this.dpasteGui.Add("Slider", "vTextTransparency NoTicks Center AltSubmit ToolTipBottom Range0-255 Invert xp yp+14 h20", "0")
        

























        
        ; Increment Controls
        this.dpasteGui.Add("GroupBox", "vincGB Section x+20 ys w120", "Adj increment")
        ; this.incUD := this.dpasteGui.Add("UpDown","-16  xs+15 ys+22 h23 vincUD Range1-" . this.contentLength - 8, "1")
        this.dpasteGui.Add("Text", "-VScroll xp49 ys+15 h40 w66 Center  vincAmnt", "1")
        this.dpasteGui["incAmnt"].SetFont("s22")
        this.dpasteGui.SetFont("s8")
        this.incUD := this.dpasteGui.Add("UpDown","Right vincUD Range1-" . this.contentLength - 8, "1")
        ; this.dpasteGui.Add("Edit", "-VScroll -WantReturn Number Center xs+5 ys+26 vincAmnt", "1")
    
    ; Add mouse wheel support for the increment control
this.dpasteGui["incAmnt"].OnEvent("Click", ObjBindMethod(this, "handleMouseWheel", "incAmnt", "incUD"))
; this.dpasteGui["incGB"].OnEvent("Click", ObjBindMethod(this, "handleMouseWheel", "incAmnt", "incUD"))



















        ; Jump Position Controls       
        this.dpasteGui.Add("GroupBox", "Section xs ys+63 w120", "Jump to position")
        ; this.jumpUD := this.dpasteGui.Add("UpDown", "xs+15 ys+23 vjumpUD -16 w40 h23 Range1-" . this.contentLength - 8, "1")
        jumpBtn := this.dpasteGui.Add("Button", "Default xp+12 ys+17 w30 h33 Center", "Go")
        this.dpasteGui.Add("Text", "-VScroll  xs68 ys+15 h40 w46 vJumpPosEdit", "1")
        this.dpasteGui["JumpPosEdit"].SetFont("s18")
        this.jumpUD := this.dpasteGui.Add("UpDown", "Right xs+15 ys+23 vjumpUD 16 w40 h23 Range1-" . this.contentLength - 8, "1")
        ; this.dpasteGui.Add("Edit", "-VScroll -WantReturn Number Right xs+5 ys+26 w45 vJumpPosEdit", "1")
        
        this.dpasteGui.SetFont("s8")
        ; Add mouse wheel support for the increment control
; this.dpasteGui["JumpPosEdit"].OnEvent("MouseMove", ObjBindMethod(this, "handleMouseWheel", "JumpPosEdit", "jumpUD"))
























        ; Event Handlers
        
        ; Setting Adjustments
        this.dpasteGui["jumpUD"].OnEvent("Change", ObjBindMethod(this, "updateJumpPreview"))
        this.dpasteGui["WindowTransparency"].OnEvent("Change", ObjBindMethod(this, "updateTransparency"))
        this.dpasteGui["TextTransparency"].OnEvent("Change", ObjBindMethod(this, "updateTextTransparency")) 
        this.incUD.OnEvent("Change", ObjBindMethod(this, "updateIncrementAmnt"))
        this.dpasteGui["JumpPosEdit"].OnEvent("Click", ObjBindMethod(this, "jumpToPosition"))
        ; this.dpasteGui["JumpPosEdit"].OnEvent("Change", ObjBindMethod(this, "updateIncrementAmnt"))
        ; this.dpasteGui["incAmnt"].OnEvent("Focus", ObjBindMethod(this, "pauseHotkeySBmsg"))
        ; this.dpasteGui["JumpPosEdit"].OnEvent("Focus", ObjBindMethod(this, "pauseHotkeySBmsg"))
        ; this.dpasteGui["incAmnt"].OnEvent("LoseFocus", ObjBindMethod(this, "resumeHotkeySBmsg"))
        
        
        ; Button Events
        jumpBtn.OnEvent("Click", ObjBindMethod(this, "jumpToPosition"))
        
    }

    
	addPreviewSection(){
        ; Add preview section
		this.dpasteGui.Add("GroupBox", "Center Section x+20 y65 w390 h140", "Preview")
		this.dpasteGui.Add("GroupBox", "Section xs+5 ys+8 w380 h50", "Adjacent elements")
		this.dpasteGui.Add("Text", "xs+10 ys+15 w360 vPreviewPrev", "Previous: ")
		this.dpasteGui.Add("Text", "xs+10 y+5 w360 vPreviewNext", "Next: ")
		; Add increment preview section
		this.dpasteGui.Add("GroupBox", "Section xs y+5 w380 h35", "Incremented Preview")
		this.dpasteGui.Add("Text", "xs+5 yp+17", "Incr. Prev.: ")
		this.dpasteGui.Add("Text", "x+ yp w135 vincPreviewPrev")
		this.dpasteGui.Add("Text", "x+1 yp", "🔷 Incr. Next: ")
		this.dpasteGui.Add("Text", "x+ yp w115 vincPreviewNext")
		
		; Add jump preview section
		this.dpasteGui.Add("GroupBox", "Section xs y+5 w380 h40", "Jump Preview (Ctrl+J to enter position)")
		this.dpasteGui.Add("Text", "xs+15 ys+17 w350 vJumpPreview")
	}
    












































    addDirectoryView() {
        this.dpasteGui.Add("GroupBox", "Section Center x+32 y65 w280 h140", "Change File (Numpad +/-)  Ctrl+O: change directory")
        dirView := this.dpasteGui.Add("TreeView", "xs+5 ys+15 w270 h90 vDirView")
        
        ; Add event handler for selection (single-click)
        dirView.OnEvent("ItemSelect", (*) => this.handleTreeViewSelection(dirView))
        
        ; Add event handler for double-click
        dirView.OnEvent("DoubleClick", (*) => this.handleTreeViewDoubleClick(dirView))
        
        ; Add "Change Directory" button
        newDirBtn := this.dpasteGui.Add("Button", "xs+8 y+5 w100 h23 Center", "Change Directory")
        this.dpasteGui.Add("Text", " vcurDirDisp Right x+8 yp+4 w150 h23", this.TVdir)
        newDirBtn.OnEvent("Click", ObjBindMethod(this, "changeDir"))
        
        ; Populate TreeView with files
        this.populateTreeView(dirView, dirPath)
        
        if RegExMatch(this.fileName, "^(.*)[\\/][^\\/]+$", &match) {
            initialDir := match[1]
            this.populateTreeView(dirView, this.fileName)
    }
    
}


    handleTreeViewSelection(treeView) {
        ; Check if an item is selected
        selectedItem := treeView.GetSelection()
        if !selectedItem {
            return  ; No item selected, do nothing
        }
        
        ; Get the text of the selected item
        selectedText := treeView.GetText(selectedItem)
        if (SubStr(selectedText, -3) != "csv") {
            return  ; Selected item is not a CSV file, do nothing
        }
        
        ; Use the directory path stored in dirPath or the parent directory of the current file
        parentFolder := RegExReplace(this.fileName, "^(.*)[\\/][^\\/]+$", "$1")
        fullPath := parentFolder . "\" . selectedText
        
        ; Load the selected file
        if FileExist(fullPath) {
            this.changeFile(fullPath)
        } else {
            MsgBox("File not found: " . fullPath)
        }
    }
    
    handleTreeViewDoubleClick(treeView) {
        ; Check if an item is selected
        selectedItem := treeView.GetSelection()
        if !selectedItem {
            return  ; No item selected, do nothing
        }
        
        ; Get the text of the selected item
        selectedText := treeView.GetText(selectedItem)
        if (SubStr(selectedText, -3) != "csv") {
            return  ; Selected item is not a CSV file, do nothing
        }
        
        ; Use the directory path stored in dirPath or the parent directory of the current file
        parentFolder := RegExReplace(this.fileName, "^(.*)[\\/][^\\/]+$", "$1")
        fullPath := parentFolder . "\" . selectedText
        
        ; Run the selected file
        if FileExist(fullPath) {
            Run(fullPath)
        } else {
            MsgBox("File not found: " . fullPath)
        }
    }
    populateTreeView(treeView, path) {
        ; Clear the tree view
        treeView.Delete()
        currentNode := ""
        if RegExMatch(path, "^(.*)[\\/][^\\/]+$", &match) {
            parentFolder := match[1]
        }
        ; Add parent node for current directory
        ; parentNode := treeView.Add(path, , "Expand")
        
        ; Find all CSV files in the directory
        Loop Files, parentFolder . "\*.csv" {
            currentNode := treeView.Add(A_LoopFileName)
            ; treeView.Add(A_LoopFileName, parentNode)
        }
        
        ; Error handling: if no files found, add a placeholder node
        if !currentNode {
            treeView.Add("No CSV files found`nTry another directory")
        }
    }












































    addUserOptions(){
        this.dpasteGui.Add("GroupBox", "Section x+28 y65 w179 h140 Right", "User options")
        defaultDirCb := this.dpasteGui.Add("CheckBox", "xs+5 ys+15 vdefaultDir", "Load current file on open")
        defaultDirCb.OnEvent("Click", (*) => this.handleOptionChange("defaultDir"))
        
        defaultPosCb := this.dpasteGui.Add("CheckBox", "xs+5 y+5 vdefaultPos", "Load current position on open")
        defaultPosCb.OnEvent("Click", (*) => this.handleOptionChange("defaultPos"))

        defaultWinPosCb := this.dpasteGui.Add("CheckBox", "xs+5 y+5 vdefaultWinPos", "Save window position")
        defaultWinPosCb.OnEvent("Click", (*) => this.handleOptionChange("defaultWinPos"))

        showSettingsCb := this.dpasteGui.Add("CheckBox", "xs+5 y+5 vshowSettings", "Hide settings on open")
        showSettingsCb.OnEvent("Click", (*) => this.handleOptionChange("showSettings"))
        
        viewSettingsCb := this.dpasteGui.Add("CheckBox", "xs+5 y+5 vloadViewSettings", "Load transparency on open")
        viewSettingsCb.OnEvent("Click", (*) => this.handleOptionChange("loadViewSettings"))
        
        viewSettingsCb := this.dpasteGui.Add("CheckBox", "xs+5 y+5 vtextScroll", "Text scroll")
        viewSettingsCb.OnEvent("Click", (*) => this.handleOptionChange("textScroll"))
        
        this.dpasteGui.Add("Text", "xs+95 yp+14", "Font Size:")
        fontSizeEdit := this.dpasteGui.Add("Edit", "x+ yp-3 Number WantReturn vdefaultSize", "8")
        this.dpasteGui.Add("UpDown", "vtextSize Range8-24")
        
        ; Add event handler for font size changes
        fontSizeEdit.OnEvent("Change", ObjBindMethod(this, "updateFontSize"))
        fontSizeEdit.OnEvent("Focus", ObjBindMethod(this, "pauseHotkeySBmsg"))
        fontSizeEdit.OnEvent("LoseFocus", ObjBindMethod(this, "resumeHotkeySBmsg"))
    }
    

    
  
    ; unfocusControl(*) {

    ; }
    
    handleOptionChange(option) {
        switch option {
            case "defaultDir":
                if (this.dpasteGui["defaultDir"].Value) {
                    ; If loading current file is enabled, also save current file path
                    IniWrite(this.curFile.name, A_ScriptFullPath . ":Stream:$DATA", "initialization", "defaultFilePath")
                }
            case "defaultPos":
                if (this.dpasteGui["defaultPos"].Value) {
                    ; If loading position is enabled, save current position
                    IniWrite(this.currentPosition, A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "lastPosition")
                }
            case "defaultWinPos":
                if (this.dpasteGui["defaultWinPos"].Value) {
                    ; If loading position is enabled, save current window position
                    WinGetPos &X, &Y,,, this.guiHwnd 
                    IniWrite("x" X " y" Y, A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "lastWinPos")
                }
            case "showSettings":
                if (this.dpasteGui["showSettings"].Value) {
                    ; If loading position is enabled, save current window position
                    IniWrite(showSettings, A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "showSettings")
                }
            case "loadViewSettings":
                if (this.dpasteGui["loadViewSettings"].Value) {
                    ; If loading position is enabled, save current window position
                    IniWrite(loadViewSettings, A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "loadViewSettings")
                }
            case "textScroll":
                if (this.dpasteGui["textScroll"].Value) {
                    ; If loading position is enabled, save current window position
                    IniWrite(textScroll, A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "textScroll")
                }
            case "textScroll":
                ; Save the state of the textScroll checkbox
                IniWrite(this.dpasteGui["textScroll"].Value, A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "textScroll")
                
                ; Start or stop the scroll timer based on the checkbox state
                if (this.dpasteGui["textScroll"].Value == 1) {
                    ; Start the scroll timer if textScroll is enabled
                    SetTimer(this.scrollTimer, 240)
                    this.moveText()
                } else {
                    ; Stop the scroll timer if textScroll is disabled
                    SetTimer(this.scrollTimer, 0)
                }
            }
            
    }
    
    showAbout(*) {
        aboutGui := Gui("+Owner" this.dpasteGui.Hwnd)
        aboutGui.Opt("+AlwaysOnTop")
        aboutGui.SetFont("s10", "Segoe UI")
        
        aboutGui.Add("GroupBox", "x10 y10 w500 h330", "Hotkeys")
        aboutGui.Add("Text", "x20 y30", "
        (
        Numpad Mult (*): Disable hotkeys
        Numpad 1-9: Paste corresponding text at cursor position
        Numpad 0: Move back one position
        Numpad Dot (.): Move forward one position
        Ctrl + Numpad 0: Decrease increment amount
        Ctrl + Numpad Dot (.): Increase increment amount
        Numpad Add (+): Move down in the directory view
        Numpad Sub (-): Move up in the directory view
        Ctrl + N: Show/hide options panel
        Ctrl + O: Open file dialog
        Ctrl + J: Open jump position dialog
        Ctrl + Numpad Add (+): Increase window opacity
        Ctrl + Numpad Sub (-): Increase window transparency
        Shift + Numpad Add (+): Make text darker
        Shift + Numpad Sub (-): Make text lighter
        Ctrl + Shift + Numpad +/-: Adjust window/text visiblity
        )")
        
        aboutGui.Add("GroupBox", "x10 y+20 w500 h190", "Options")
        aboutGui.Add("Text", "x20 yp+20 w480", "
        (
        • Window and Text Transparency can be adjusted with sliders
        • Increment amount changes how many positions you move with Numpad 0/Dot
        • Jump to position lets you quickly navigate to a specific element
        • Files can be loaded by selecting them in the directory view
        • All csv files in parent directory will appear in viewer
        • Preview shows surrounding elements for context
        • The mouse wheel adjusts increment and jump controls when mouse hovers over them
        )")
        
        aboutGui.Add("Button", "x220 y+15 w80 h25", "OK").OnEvent("Click", (*) => aboutGui.Destroy())
        
        aboutGui.Show("")
    }
        
        
        
        
        
        
        
        
        
    addStatusBar() {
        this.statusBar := this.dpasteGui.Add("StatusBar",, "  Current File: ")
        this.statusBar.SetParts(70, 170, 390)
        this.updateStatusBar()  ; Initial update
    }
    


    updateStatusBar() {
        formatedFileName := ""
        fileNameMaxLength := 33
        ; Calculate current line based on position and elements per line

        elementCount := 1
        totalLines := this.totalLines
        curPosition := this.currentPosition + 1
        lineCounterArr := this.lineCounterArr
        currentLine := ""
        
        Loop totalLines 
        {
        ; msgbox lineCounterArr[A_Index]
            currentLine := A_Index
            if curPosition <= lineCounterArr[A_Index] {
                break
            }
        }
        
        ; Calculate current elements range
        startElement := this.currentPosition + 1
        endElement := Min(this.currentPosition + 9, this.curFile.contentArr.Length)
        formatedFileName := this.curFile.name 
        ; Update status bar text
        if (StrLen(formatedFileName) > fileNameMaxLength)
            {
            formatedFileName :=  SubStr(this.curFile.name , StrLen(this.curFile.name ) * .15)
            Loop
                formatedFileName :=  SubStr(formatedFileName, (StrLen(formatedFileName) * .12))
            Until StrLen(formatedFileName) < fileNameMaxLength 
            formatedFileName := "..." . formatedFileName
        }
        this.statusBar.SetText(formatedFileName,2)
        this.statusBar.SetText("`tScript Position: " . this.currentPosition + 1 . ". Elements " startElement "-" endElement " out of " this.curFile.contentArr.Length " loaded. Line " . currentLine . " out of " totalLines, 3)
        this.statusBar.SetText("`t`tNum+* Disable Triggers - Ctrl+Numpad0/Dot Adjust Increment Amount - Ctrl+N Hide Settings",4)
    }












































    increaseTextTransparency(){
        this.dpasteGui["TextTransparency"].Value += 30
        this.updateTextTransparency
    }
    decreaseTextTransparency(){
        this.dpasteGui["TextTransparency"].Value -= 30
        this.updateTextTransparency
    }
    
    increaseTransparency(){
        this.dpasteGui["WindowTransparency"].Value += 25
        this.updateTransparency
        }
        
    decreaseTransparency(){
        this.dpasteGui["WindowTransparency"].Value -= 25
        this.updateTransparency
    }
    
    updateIncrementAmnt(*){
        this.incAmount := (this.dpasteGui["incAmnt"].Value != "" && this.dpasteGui["incAmnt"].Value > 0)   ? this.dpasteGui["incAmnt"].Value : 1
    }
    
    updateTransparency(*) {
        winTrans := this.dpasteGui["WindowTransparency"].Value
        WinSetTransparent(winTrans, this.dpasteGui)
    }
    

	updateTextTransparency(*) {
		textTrans := Format("{:X}", this.dpasteGui["TextTransparency"].Value)
		Loop 9 {
			this.dpasteGui["Text" A_Index].Opt("c" textTrans . textTrans . textTrans)
		}
	}
	
	
    moveText(*) {
        ; Check if textScroll is enabled
        if (this.dpasteGui["textScroll"].Value == 1) {
            for controlName in this.scrollArray {
                ctrlName := controlName
                currentText := this.dpasteGui[controlName].Text
                if (StrLen(currentText) > this.textLengthThreshold) {
                    ; Check if the last character is a space
                    leftText := SubStr(currentText, 2)
                    rightText := SubStr(currentText, 1, 1)
                    newText := leftText . rightText
                    
                    this.dpasteGui[ctrlName].Value := newText
                }
            }
        }
    }
    
    ; moveText(*) {
    ;     if this.dPasteGui[textScroll].Value
    ;         return
    ;     for controlName in this.scrollArray {
    ;         ctrlName := controlName
    ;         currentText := this.dpasteGui[controlName].Text
    ;         if (StrLen(currentText) > this.textLengthThreshold) {
    ;             ; Check if the last character is a space
    ;             leftText := SubStr(currentText, 2)
    ;             rightText := SubStr(currentText, 1, 1)
    ;             newText := leftText . rightText
                
    ;             this.dpasteGui[ctrlName].Value := newText
    ;         }
    ;     }
    ; }
    
    
    ; GuiControlGet, loopText,, %loopTarget%
    ; loopText_Len := StrLen(loopText) - 1
    ; leftText := SubStr(loopText,2,loopText_Len)
    ; rightText := SubStr(loopText,1,1)
    ; GuiControl, Text, %loopTarget%, %leftText%%rightText%
    ; sleep, 10
    ; }






    settingsToggle() {
        this.currentToggle := !this.currentToggle
        if (!this.currentToggle) {
            this.dpasteGui.Move(,,, 140)
        } else {
            this.dpasteGui.Move(,,, 308)
        }
    }












































    showFontSizeDialog() {
        fontSizeGui := Gui("+Owner" this.dpasteGui.Hwnd)
        fontSizeGui.OnEvent("Escape", (*) => fontSizeGui.Destroy())
        fontSizeGui.OnEvent("Close", (*) => fontSizeGui.Destroy())
        
        fontSizeGui.SetFont("s10", "Segoe UI")
        fontSizeGui.Add("Text",, "Font Size (8-40):")
        
        fontSizeEdit := fontSizeGui.Add("Edit", "Number vFontSize w60", this.dpasteGui["defaultSize"].Value)
        fontSizeGui.Add("UpDown", "Range8-40", this.dpasteGui["defaultSize"].Value)
        
        applyBtn := fontSizeGui.Add("Button", "Default w80", "Apply")
        applyBtn.OnEvent("Click", (*) => (
            this.dpasteGui["defaultSize"].Value := fontSizeEdit.Value,
            this.updateFontSize(),
            fontSizeGui.Destroy()
        ))
        cancelBtn := fontSizeGui.Add("Button", "x+10 w80", "Cancel")
        cancelBtn.OnEvent("Click", (*) => fontSizeGui.Destroy())
        
        fontSizeGui.Show()
    }
    
    








































    showJumpPosDialog() {
        jumpPosGui := Gui("+Owner" this.dpasteGui.Hwnd)
        jumpPosGui.OnEvent("Escape", (*) => jumpPosGui.Destroy())
        jumpPosGui.OnEvent("Close", (*) => jumpPosGui.Destroy())
        
        jumpPosGui.SetFont("s10", "Segoe UI")
        jumpPosGui.Add("Text","y22", "Jump to... (1-" . this.contentLength - 8 . "):")
        
        ; Add the Edit and UpDown controls
        jumpPosEdit := jumpPosGui.Add("Edit", "Number yp-3 x+11 vJumpMenuEdit w60", this.dpasteGui["JumpPosEdit"].Value)
        jumpPosGui.Add("UpDown", "Range1-" . this.contentLength - 8, this.dpasteGui["jumpUD"].Value)
        
        ; Add preview Text control
        jumpMenuPreview := jumpPosGui.Add("Text", "xs+40 y+20 w480 h53 vJumpMenuPreview", "Preview will show here")
        ; jumpPosEdit.OnEvent("Change", () => MsgBox jumpPosEdit.Value)
        
        ;    OnEvent for jumpMenuEdit to update the preview
        jumpPosEdit.OnEvent("Change", (*) => (
            this.dpasteGui["JumpPosEdit"].Value := jumpPosEdit.Value
            
            jumpMenuPreview.Value := this.curFile.contentArr[jumpPosEdit.Value]  ; Call the method to update preview
            ; jumpMenuPreview.Value := this.dpasteGui["JumpPreview"].Value  ; Call the method to update preview
            ; jumpPosMenuPreview.Value := this.updateJumpPreview()  ; Call the method to update preview
            ; jumpPosPreview.Value := this.updateJumpPreview()  ; Call the method to update preview
        ))
        
        
        ; Add buttons
        applyBtn := jumpPosGui.Add("Button", "Default x+10 y+15 w80", "Apply")
        applyBtn.OnEvent("Click", (*) => (
            this.dpasteGui["JumpPosEdit"].Value := jumpPosEdit.Value
            this.jumpToPosition()
            jumpPosGui.Destroy()
        ))
        
        cancelBtn := jumpPosGui.Add("Button", "x+10 w80", "Cancel")
        cancelBtn.OnEvent("Click", (*) => jumpPosGui.Destroy())
        
        jumpPosGui.Show()
    }
    










    handleMouseWheel(controlName, upDownName, guiCtrl, info) {
        ; Check if the mouse wheel is being scrolled
        if (info = "MouseWheelUp" || info = "MouseWheelDown") {
            ; Get the current value of the UpDown control
            currentValue := this.dpasteGui[upDownName].Value
            
            ; Adjust the value based on the mouse wheel direction
            if (info = "MouseWheelUp") {
                newValue := currentValue + 1
            } else if (info = "MouseWheelDown") {
                newValue := currentValue - 1
            }
            
            ; Ensure the new value is within the valid range
            minValue := this.dpasteGui[upDownName].Min
            maxValue := this.dpasteGui[upDownName].Max
            newValue := Min(Max(newValue, minValue), maxValue)
            
            ; Update the UpDown control and the associated text control
            this.dpasteGui[upDownName].Value := newValue
            this.dpasteGui[controlName].Value := newValue
            
            ; Trigger any associated logic (e.g., update the display)
            this.updateIncrementAmnt()
        }
    }
































    updateFontSize(*) {
        newSize := this.dpasteGui["defaultSize"].Value
        if (newSize = "" || newSize < 8 || newSize > 24) {
            ToolTip("Font size must be between 8 and 24")
            SetTimer () => ToolTip(), -3000
            return
        }
        newPos := 26    ; default y position
        newHeight := 15 ; default height
        
        switch 
        {
            ; Smallest text sizes
            case newSize >= 8 && newSize <= 10: 
                newPos := newPos
                newHeight := newHeight
            
            ; Small-medium text sizes
            case newSize > 10 && newSize <= 13:
                newPos := 25
                newHeight := 22
                
            ; Medium text sizes
            case newSize > 13 && newSize <= 16:
                newPos := 20
                newHeight := 28
                this.textLengthThreshold := 9
                
            ; Medium text sizes2
            case newSize > 16 && newSize <= 18:
                newPos := 18
                newHeight := 29
                this.textLengthThreshold := 9
                
            ; Medium-large text sizes
            case newSize > 18 && newSize <= 23:
                newPos := 15
                newHeight := 35
                this.textLengthThreshold := 7
            ; Large text sizes
            case newSize >= 24:
                newPos := 14
                newHeight := 37
                this.textLengthThreshold := 4
        }
        
        ; Update all text controls
        Loop 9 {
            this.dpasteGui["Text" A_Index].SetFont("s" newSize)
            this.dpasteGui["Text" A_Index].Move(,newPos,,newHeight)
        }
        this.updateDisplay
    }












































    updateJumpPreview(*)
    {
        curContentLength := this.curFile.contentArr.Length 
        if this.jumpUD.Value == ""
            return
        previewText := ""
        jumpPos := Integer(this.jumpUD.Value)
        if (jumpPos < 1 || jumpPos > curContentLength - 8 || jumpPos == "") 
        {
            previewText := "N/A - Out of bounds! Pressing 'Go' jumps to closest bound"
        }
        else
        {
            Loop this.previewSize
            {
                if (jumpPos + A_Index - 1 <= this.curFile.contentArr.Length) 
                {
                    if A_Index < this.previewSize
                    {
                        previewText .= this.curFile.contentArr[jumpPos + A_Index - 1] . " -- "
                    }
                    else previewText .= this.curFile.contentArr[jumpPos + A_Index - 1] . "..."
                }
            }
        }
        this.dpasteGui["JumpPreview"].Value := previewText
    }














































    jumpToPosition(*) 
    {
        position := this.jumpUD.Value
        if position == ""
            position := 0
        position := Integer(position)
        switch
        {
            ; if (position >= 1 && position <= this.curFile.contentArr.Length - 8) {
            case (position < 1):
            {
                jumpPos := 0
                jumpEditVal := 1
            }
            case (position > this.curFile.contentArr.Length - 8):
            {
                jumpPos := this.curFile.contentArr.Length - 9
                jumpEditVal := jumpPos + 1
            }
            Default:
            {
                jumpPos := position - 1
                jumpEditVal := position
            }
        }
        this.currentPosition := jumpPos
    	this.dpasteGui["JumpPosEdit"].Value := jumpEditVal
        this.updateDisplay()
        this.updateJumpPreview()
    }












































    updateDisplay() {
        this.scrollArray := []  ; Reset scroll array
        ; incEdit := this.dpasteGui["mcstls_updown321"]
        this.incUD.Opt("Range1-" this.contentLength)
        this.jumpUD.Opt("Range1-" this.contentLength)
        ; this.dpasteGui.Opt("UpDown","Range1-" . this.curFile.contentArr.Length - 8, "1")
        Loop 9 {
            currentArrayPos := this.currentPosition + A_Index
            if (currentArrayPos <= this.curFile.contentArr.Length) {
                textContent := this.curFile.contentArr[currentArrayPos]
                this.dpasteGui["Text" A_Index].Value := textContent . "   "
                
                ; Add to scroll array if text is too long
                if (StrLen(textContent) > this.textLengthThreshold) {
                    this.scrollArray.Push("Text" A_Index)
                }
            }
        }
        
        ; Update preview text
        incAmount := this.incAmount
        
        ; Adjacent Previous
        prevText := "Previous: "
        prevDelim := " -- "
        prevPosStart := this.currentPosition 
        prevPos := prevPosStart
        if (prevPos > 0)
        {
            Loop 
                {
                    prevText .= this.curFile.contentArr[prevPos]
                    prevPos -= 1
                    if (prevPos == 0 || prevPosStart - prevPos == this.previewSize)
                        break
                    prevText .= prevDelim
                }        
        }
        else
            prevText := "Previous: N/A - at file start"
        
        
        ; Adjacent Next
        nextText := "Next: "
        nextDelim := " -- "
        nextPosStart := this.currentPosition + 9 + 1
        nextPos := nextPosStart
        if (nextPos <= this.curFile.contentArr.Length)
        {
            ; msgbox "nextPos= " . nextPos . "--Start-Pos=" nextPos - nextPosStart
            Loop
                {
                    nextText .= this.curFile.contentArr[nextPos]
                    nextPos += 1
                    ; msgbox "nextPos= " . nextPos 
                if (nextPos > this.curFile.contentArr.Length || nextPos - nextPosStart >= this.previewSize)
                    break
                nextText .= prevDelim
            }
        }
        else 
            nextText := "Next: N/A - at end of file"
            
        ; Increment Previous
        incPrevPos := Max(1, this.currentPosition - incAmount + 1)
        incPrevText := this.curFile.contentArr[incPrevPos]
        if incPrevPos == 1
            incPrevText := "N/A - at file start"
        else (StrLen(incPrevText) > this.optionTextLengthThreshold)
            this.scrollArray.Push("incPreviewPrev")
        
        ; Increment Next
        incNextPos := Min(this.curFile.contentArr.Length, this.currentPosition + 9 + incAmount)
        incNextText := this.curFile.contentArr[incNextPos]
        if incNextPos + 1 > this.curFile.contentArr.Length
            incNextText := "N/A - beyond end of file"
        if (StrLen(incNextText) > this.optionTextLengthThreshold)
            this.scrollArray.Push("incPreviewNext")
    

        this.TVdir := ""
        if RegExMatch(this.fileName, "[\\/](?<GrandParent>[^\\/]+)[\\/][^\\/]+[\\/][^\\/]+$", &match) {
            this.grandParentDir := match[1]
        }
        if RegExMatch(this.fileName, "[\\/](?<Parent>[^\\/]+)[\\/][^\\/]+$", &match) {
            this.parentDir := match[1]
        }
        ; RegExReplace(this.fileName, "^(.*)[\\/][^\\/]+$", this.parentDir)
        ; ; this.parentDir := RegExReplace(this.fileName, "^(.*)[\\/][^\\/]+[\\/]?$", "")
        this.TVdir := "\" this.grandParentDir . "\" . this.parentDir . "\"




        
        this.dpasteGui["incPreviewPrev"].Value := incPrevText
        this.dpasteGui["incPreviewNext"].Value := incNextText
        this.dpasteGui["PreviewPrev"].Value := prevText
        this.dpasteGui["PreviewNext"].Value := nextText
        this.dpasteGui["incAmnt"].Value := this.incAmount
        this.dpasteGui["incUD"].Value := this.incAmount
        ; this.dpasteGui["JumpPosEdit"].Value := this.currentPosition
        ; this.dpasteGui["jumpUD"].Value := this.currentPosition
        this.dpasteGui["curDirDisp"].Value :=  this.TVdir
        
        this.updateStatusBar() 
    }











































    printElement(contentArray) {
        sleepamnt := 50
        old_clip := ClipboardAll()
        Sleep sleepamnt
        A_Clipboard := contentArray[Integer(SubStr(A_ThisHotkey, 7)) + this.currentPosition]
        Send "^v"
        Sleep sleepamnt
        A_Clipboard := old_clip
        }
    
    
    
    
    
    
    

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ; change file and reset variables
    changeFile(filePath) {
        this.curFile := {}
        if !filePath
            filePath := FileSelect()
        this.currentPosition := 0
        this.curFile := readTargetFile(filePath)
        this.fileName := this.curFile.name
        this.contentLength := this.curFile.contentArr.Length
        this.lineCounterArr := this.curFile.lineArr
        this.totalLines := this.curFile.lines
        
        
        
        
        
        ;refresh display
        this.updateDisplay()
        this.updateJumpPreview()
    }





    








    showEmptyStringDialog() {
        emptyStringGui := Gui("+Owner" this.dpasteGui.Hwnd)
        emptyStringGui.OnEvent("Escape", (*) => emptyStringGui.Destroy())
        emptyStringGui.OnEvent("Close", (*) => emptyStringGui.Destroy())
        
        emptyStringGui.SetFont("s10", "Segoe UI")
        emptyStringGui.Add("Text",, "Reset string for empty fields:")
        
        emtpyStringValue := emptyStringGui.Add("Edit", "vdefaultEmptyString w100", IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "defaultEmptyString", "¯\_(ツ)_/¯"))
        
        applyBtn := emptyStringGui.Add("Button", "Default w80", "Save")
        applyBtn.OnEvent("Click", (*) => (
            IniWrite(emtpyStringValue.value,  A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "defaultEmptyString")
            this.changeFile(this.fileName),
            emptyStringGui.Destroy()
        ))
        cancelBtn := emptyStringGui.Add("Button", "x+10 w80", "Cancel")
        cancelBtn.OnEvent("Click", (*) => emptyStringGui.Destroy())
        
        emptyStringGui.Show()
    }






























    changeDir(*) {
        treeView := this.dpasteGui["DirView"]
    
        ; Open a file selection dialog to choose a new directory
        newFilePath := FileSelect(3,, "Select a CSV file")
        if newFilePath == ""  ; If the user cancels, do nothing
            return
    changingFile := readTargetFile(newFilePath)
        ; Extract the parent directory from the selected file path
        if RegExMatch(newFilePath, "^(?<Drive>[A-Za-z]:)?(?<Path>.*)[\\/][^\\/]+$", &match) {
            ; Extract the drive and path
            newDrive := match["Drive"]
            newPath := match["Path"]
    ; msgbox newDrive
    ; msgbox newPath
            ; Update the directory tracking variables
            this.parentDir := RegExReplace(newPath, "^.*[\\/]([^\\/]+)$", "$1")
            this.grandParentDir := RegExReplace(newPath, "^.*[\\/]([^\\/]+)[\\/][^\\/]+$", "$1")
            this.TVdir := "\" . this.grandParentDir . "\" . this.parentDir . "\"
            this.fileName := newFilePath
            ; Update the directory text control in the GUI
            this.dpasteGui["DirView"].Text := this.TVdir
            this.changeFile(newFilePath)
            ; Populate the TreeView with CSV files from the new directory
            this.populateTreeView(treeView, newFilePath)
    
            ; Update the display and jump preview
            this.updateDisplay()
            this.updateJumpPreview()
        }
    }

    ; changeDir(*) {
    ;     treeView := this.dpasteGui["DirView"]

    ;     newdirPath := FileSelect()
    ;     if newdirPath == ""
    ;         return
    ;     this.populateTreeView(treeView,newdirPath)
    ;     this.updateDisplay()
    ;     this.updateJumpPreview()
    ; }
    
    
    
    loadSelectedFile(*) {
        this.curFile := {}
        
        ; Check if an item is selected in the TreeView
        selectedItem := this.dpasteGui["DirView"].GetSelection()
        if !selectedItem {
            MsgBox("No file selected. Please select a file from the directory view.")
            return
        }
        
        ; Get the text of the selected item
        selectedText := this.dpasteGui["DirView"].GetText(selectedItem)
        if (SubStr(selectedText, -3) != "csv") {
            MsgBox("Selected item is not a CSV file.")
            return
        }
        
        ; Use the directory path stored in dirPath or the parent directory of the current file
        parentFolder := RegExReplace(this.fileName, "^(.*)[\\/][^\\/]+$", "$1")
        fullPath := parentFolder . "\" . selectedText
        
        if FileExist(fullPath) {
            this.changeFile(fullPath)
        } else {
            MsgBox("File not found: " . fullPath)
        }
    }
    
    runSelectedFile(*) {
        ; Check if an item is selected in the TreeView
        selectedItem := this.dpasteGui["DirView"].GetSelection()
        if !selectedItem {
            MsgBox("No file selected. Please select a file from the directory view.")
            return
        }
        
        ; Get the text of the selected item
        selectedText := this.dpasteGui["DirView"].GetText(selectedItem)
        if (SubStr(selectedText, -3) != "csv") {
            MsgBox("Selected item is not a CSV file.")
            return
        }
        
        ; Use the directory path stored in dirPath or the parent directory of the current file
        parentFolder := RegExReplace(this.fileName, "^(.*)[\\/][^\\/]+$", "$1")
        fullPath := parentFolder . "\" . selectedText
        
        if FileExist(fullPath) {
            Run(fullPath)
        } else {
            MsgBox("File not found: " . fullPath)
        }
    }
    
    
    
    
    ; loadSelectedFile(*) {
    ;     this.curFile := {}
        
    ;     try selectedItem := this.dpasteGui["DirView"].GetSelection()  ; Get selected item
    ;     selectedText := this.dpasteGui["DirView"].GetText(selectedItem)
    ;     if (SubStr(selectedText,-3) != "csv")
    ;         return
    ;     fullPath := (A_WorkingDir . "\" . selectedText)
    ;     this.changeFile(fullPath)
    ;     }
        
    ; runSelectedFile(*) {
    ;     selectedItem := this.dpasteGui["DirView"].GetSelection()  ; Get selected item
    ;     selectedText := this.dpasteGui["DirView"].GetText(selectedItem)
    ;     if (SubStr(selectedText,-3) != "csv")
    ;         return
    ;     fullPath := (A_WorkingDir . "\" . selectedText)
    ;    run fullPath
    ;     }
    











































    saveSettings() {
        iniPath := A_ScriptFullPath . ":Stream:$DATA"
        
        ; Save initialization settings
        IniWrite(this.curFile.name, iniPath, "initialization", "defaultFilePath")
        IniWrite(dirPath, iniPath, "initialization", "defaultDirPath")
        
        ; Save user settings
        IniWrite(this.dpasteGui["defaultSize"].Value, iniPath, "UserSettings", "fontSize")
        IniWrite(this.dpasteGui["defaultDir"].Value, iniPath, "UserSettings", "loadLastFile")
        IniWrite(this.dpasteGui["defaultPos"].Value, iniPath, "UserSettings", "loadLastPosition")
        IniWrite(this.dpasteGui["defaultWinPos"].Value, iniPath, "UserSettings", "loadLastWinPos")
        IniWrite(this.dpasteGui["showSettings"].Value, iniPath, "UserSettings", "showSettings")
        IniWrite(this.dpasteGui["loadViewSettings"].Value, iniPath, "UserSettings", "loadViewSettings")
        IniWrite(this.dpasteGui["textScroll"].Value, iniPath, "UserSettings", "textScroll")
        IniWrite(this.currentPosition, iniPath, "UserSettings", "lastPosition")
        IniWrite(this.dpasteGui["WindowTransparency"].Value, iniPath, "UserSettings", "windowTransparency")
        IniWrite(this.dpasteGui["TextTransparency"].Value, iniPath, "UserSettings", "textTransparency")
        curWinSet := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "loadLastWinPos", "0")
        curPosSet := IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "loadLastPosition", "0")
        WinGetPos &X, &Y,,, this.guiHwnd 
        IniWrite("x" X " y" Y, iniPath, "UserSettings", "lastWinPos")
    }
    
    loadSettings() {
        ; Set font size
        this.dpasteGui["defaultSize"].Value := defaultFontSize
        this.updateFontSize()
        
        ; Set checkbox states
        this.dpasteGui["defaultDir"].Value := loadLastFile
        this.dpasteGui["defaultPos"].Value := loadLastPos
        this.dpasteGui["defaultWinPos"].Value := loadLastWinPos
        this.dpasteGui["showSettings"].Value := showSettings
        this.dpasteGui["loadViewSettings"].Value := loadViewSettings
        this.dpasteGui["textScroll"].Value := textScroll
    
        ; Initialize the scroll timer based on the textScroll checkbox state
        if (this.dpasteGui["textScroll"].Value == 1) {
            SetTimer(this.scrollTimer, 240)
        } else {
            SetTimer(this.scrollTimer, 0)
        }
        
        ; Set position if enabled
        if (loadLastPos == "1") {
            this.currentPosition := Integer(lastPosition)
            this.updateDisplay()
        }
        
        ; Set transparency values
        if (loadViewSettings == "1")
        {
            this.dpasteGui["WindowTransparency"].Value := windowTransparency
            this.updateTransparency()
            this.dpasteGui["TextTransparency"].Value := textTransparency
            this.updateTextTransparency()
        }
        
        
        if (this.dpasteGui["defaultDir"].Value) 
        {
            savedPath := IniRead(A_ScriptFullPath . ":Stream:$DATA", "initialization", "defaultFilePath", A_ScriptDir . "\Ltest.csv")
            if (FileExist(savedPath))
            {
                this.changeFile(savedPath)
            }
        }  
        
        if (this.dpasteGui["defaultPos"].Value) {
            savedPos := Integer(IniRead(A_ScriptFullPath . ":Stream:$DATA", "UserSettings", "lastPosition", "0"))
            if (savedPos >= 0 && savedPos < this.curFile.contentArr.Length) {
                this.currentPosition := savedPos
                this.updateDisplay()
            }
        }
        this.updateDisplay()
          
    }   
    handleExit(ExitReason, ExitCode) {
        SetTimer(this.scrollTimer, 0)  ; Stop the scroll timer
        ; Save all settings before exit
        this.saveSettings()
        
        return 0
    }












































    navigateTreeView(direction) {
        treeView := this.dpasteGui["DirView"]
        currentSelection := treeView.GetSelection()
        
        ; If no selection, start at the first or last item
        if (!currentSelection) {
            if (direction > 0) {
                ; Move to the first item
                firstItem := treeView.GetChild(0)
                if (firstItem) {
                    treeView.Modify(firstItem, "Select")  ; Select the first item
                }
            } else {
                ; Move to the last item
                lastItem := treeView.GetPrev(0)  ; Start from the root and get the last item
                if (lastItem) {
                    treeView.Modify(lastItem, "Select")  ; Select the last item
                }
            }
            return
        }
        
        ; Move up or down based on the direction
        if (direction > 0) {
            ; Move down to the next item
            nextItem := treeView.GetNext(currentSelection)
            if (nextItem) {
                treeView.Modify(nextItem, "Select")  ; Select the next item
            }
        } else {
            ; Move up to the previous item
            prevItem := treeView.GetPrev(currentSelection)
            if (prevItem) {
                treeView.Modify(prevItem, "Select")  ; Select the previous item
            }
        }
        
        ; Optionally, load the selected file automatically
        this.loadSelectedFile()
    }
    
    exitApp() {    
        ; Force script termination
        ExitApp 0
    }
    
}

; Create global instance of GUI
global MainGui := drPaster()
guiHwnd := MainGui.guiHwnd

MainGui.updateDisplay()

; Hotkey handlers
#SuspendExempt
; Reload Script
NumpadMult::
{
  Suspend
  if A_IsSuspended
    MainGui.statusBar.SetText("`t****SCRIPT SUSPENDED****TRIGGERS DISABLED****PRESS Numpad+Mult TO RESUME****",4)
  else
    MainGui.statusBar.SetText("`t`tNum+* Disable Triggers - Ctrl+Numpad0/Dot Adjust Increment Amount - Ctrl+N Hide Settings",4)
}
#SuspendExempt false


#HotIf !GetKeyState("Shift", "P")
^NumpadAdd:: {
    MainGui.increaseTransparency()
}

^NumpadSub:: {
    MainGui.decreaseTransparency()
}
#HotIf

#HotIf !GetKeyState("Ctrl", "P")
+NumpadAdd:: {
    MainGui.decreaseTextTransparency()
}

+NumpadSub:: {
    MainGui.increaseTextTransparency()
}
#HotIf

^+NumpadSub:: {
    MainGui.increaseTextTransparency()
    MainGui.decreaseTransparency()
}

^+NumpadAdd:: {
    MainGui.decreaseTextTransparency()
    MainGui.increaseTransparency()
}

^NumpadDot:: {
    MainGui.incAmount += 1
    if (MainGui.incAmount > MainGui.curFile.contentArr.Length) {
        MainGui.incAmount := 1
        ToolTip("Increment amount exceeds length of file! Resetting to 1")
        SetTimer(() => ToolTip(), -5000)
    }
    MainGui.updateDisplay()
}

^Numpad0:: {
    MainGui.incAmount -= 1
    if (MainGui.incAmount < 1) {
        MainGui.incAmount := 1
        ToolTip("Increment amount minimum reached")
        SetTimer(() => ToolTip(), -5000)
    }
    MainGui.updateDisplay()
}

NumpadAdd:: {
    MainGui.navigateTreeView(1)
}

NumpadSub:: {
    MainGui.navigateTreeView(-1)
}

NumpadDot:: {
    MainGui.currentPosition := Min(MainGui.curFile.contentArr.Length - 9, MainGui.currentPosition + MainGui.incAmount)
    MainGui.updateDisplay()
}

Numpad0:: {
    ; If WinActive(guiHwnd) {
    ;     focusedHwnd := ControlGetFocus("A")
    ;     FocusedClassNN := ControlGetClassNN(FocusedHwnd)
    ;     ; If (InStr(FocusedClassNN, "Edit")) {
    ;     ;     MainGui.statusBar.SetText("`t****EDIT CONTROL FOCUS DETECTED****SUSPENDING HOTKEY TRIGGERS****", 3)
    ;     ;     send SubStr(A_ThisHotkey, 7)
    ;     ;     return
    ;     ; }
    ; }
    MainGui.currentPosition := Max(0, MainGui.currentPosition - MainGui.incAmount)
    MainGui.updateDisplay()
}

; Number pad handlers for printing elements
Numpad1::
Numpad2::
Numpad3::
Numpad4::
Numpad5::
Numpad6::
Numpad7::
Numpad8::
Numpad9::
{
    ; If WinActive(guiHwnd) {
    ;     focusedHwnd := ControlGetFocus("A")
    ;     FocusedClassNN := ControlGetClassNN(FocusedHwnd)
    ;     If InStr(FocusedClassNN, "Edit") {
    ;         MainGui.statusBar.SetText("`t****EDIT CONTROL FOCUS DETECTED****SUSPENDING HOTKEY TRIGGERS****", 3)
    ;         send SubStr(A_ThisHotkey, 7)
    ;         return
    ;     }
    ; }
    content := MainGui.curFile.contentArr
    MainGui.printElement(content)
}
#HotIf