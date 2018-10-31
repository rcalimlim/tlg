;///////////////////////////////////////////////////////////////////////////////
; SCRIPT SETTINGS
;///////////////////////////////////////////////////////////////////////////////
#SingleInstance force
#NoEnv ; Recommended for performance and compatibility with future 
       ; AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior 
               ; speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.

;///////////////////////////////////////////////////////////////////////////////
; TLG ROSS v3.5
;///////////////////////////////////////////////////////////////////////////////
; This script translates shorthand TLG information entered by the user into
; proper delorean codes. This script, in general, defaults to silent errors
; and returns nothing due to usability concerns when keying information in 
; quick succession.
; 
; For now, shorthand customization and project entries must be entered into the
; script itself. Future enhancements look to importing that information from a
; a spreadsheet for ease-of-maintenance.
; 
; This script also defaults to sending the translated inputs directly after
; the user enters the TLG information, but can be turned off.
;
; Update: 2018 October 30

;///////////////////////////////////////////////////////////////////////////////
; DEFINE GLOBALS
;///////////////////////////////////////////////////////////////////////////////
global __all__maintable := make_table("Main")
global __all__desctable := make_table("TLG Descriptions")

;///////////////////////////////////////////////////////////////////////////////
; DEFINE HOTKEYS 
;///////////////////////////////////////////////////////////////////////////////
; Run Script
; Shift + Alt + J
+!j::
msgbox % tlg_wrapper()
return

; Reload Script
; Shift + Alt + S
+!s::
reload
return

;///////////////////////////////////////////////////////////////////////////////
; DEFINE FUNCTIONS
;///////////////////////////////////////////////////////////////////////////////
; Name:         get_input
; Description:  Prompts and returns a user's input.
; Parameters:   None
; Called by:    format_inputs
; Returns:      string (if not cancelled)
;               -1 (if cancelled)
get_input() {
    msg = [org] [tlg], [desc]
    inputbox, str, TLG Ross, %msg%,, 200, 150 ; inputbox size 200x150
    if (errorlevel != 0) {  ; return ErrorLevel integer if cancelled
        return -1
    }
    else return str ; otherwise return the string input
}
;///////////////////////////////////////////////////////////////////////////////
; Name:         str_to_arr
; Description:  Converts string to array using passed delimier and omits passed
;               characters.
; Parameters:   str: string to create array
;               delim: delimiter string (defaults to nothing)
;               omit: characters to exclude from strings (defaults to nothing)
; Called by:    format_inputs
; Returns:      arr: array created from string
;               -1 (bad input)
str_to_arr(str, delim:="", omit:="") {
    if (!str || str == -1) {
        return -1
    }
    else {
        return arr := strsplit(str, delim, omit)
    }
}
;///////////////////////////////////////////////////////////////////////////////
; Name:         format_inputs
; Description:  Calls get_input and returns two arrays. Array 1 contains TLG
;               information before the comma, Array 2 contains the description
;               override if any.
; Parameters:   byref variables passed out by name when called
; Called by:    tlg_wrapper
; Returns:      tlgarr: array from string delimited by spaces, excluding commas
;               descrip: array containing all information after a comma, if any
format_inputs(byref tlgarr, byref descrip) {
    userinput := get_input()
    if (userinput == -1 or userinput == "")
        tlgarr := descrip := -1
    else {
        tlgarr := str_to_arr(userinput, " ", ",") ; array
        descrip := str_to_arr(userinput, ",")[2] ; string
    }
}
;///////////////////////////////////////////////////////////////////////////////
; Name:         num_to_alpha
; Description:  Takes an integer and returns its alphabetic equivalent. Errors
;               passed value is not an integer or not within 1-26.
; Parameters:   int: integer to convert to alpha character
; Called by:    get_excel_col
; Returns:      alphabetic character (good input)
;               error message (bad input)
num_to_alpha(int) {
    alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if int is not integer
        return "Non-integer input"
    else if (int < 0 || !int || int > 27)
        return "Integer out of alphabetic bounds"
    else return substr(alphabet, int, 1)
}
;///////////////////////////////////////////////////////////////////////////////
; Name:         get_excel_col
; Description:  I cannot for the life of me figure out how VBA works, so I had
;               to write a function that converts the numeric column returned
;               from a SpecialCell lookup into familiar alphabetic excel
;               notation. 
;               This is a recursive function.
; Parameters:   column_num: Excel numeric column ID
;               divisor: modulo divisor (should always be 26 but whatever)
; Called by:    get_excel_col (recursively)
;               make_table
; Returns:      alphabetic translation of col ID (good input)
;               error message (bad input)
get_excel_col(column_num) {
    errormsg := "Parameters must be positive integers"
    if column_num is not integer
        return % errormsg
    else if (column_num <= 0)
        return % errormsg
    else if (column_num <= 26)
        return % num_to_alpha(column_num)
    else {
        remainder := mod(column_num, 26)
        column_num := floor(column_num/26)
        return % get_excel_col(column_num) . num_to_alpha(remainder)
    }
}
;///////////////////////////////////////////////////////////////////////////////
; Name:         make_table
; Description:  Gets an excel workbook from passed file path, and returns an
;               array object for passed sheet.
; Parameters:   sheet: sheet name
;               file_path: file path of excel workbook, defaults to Ross'
; Called by:    __all__maintable (global)
;               __all__desctable (global)
; Returns:      array object
make_table(sheet, file_path := "C:\Users\Ross\Desktop\matrix.xlsx") {
    oWorkbook := comobjget(file_path)
    ; VBA crap probably
    lastrow := oWorkbook.Sheets(sheet).Range("A:A").SpecialCells(11).Row
    lastcol := oWorkbook.Sheets(sheet).Range("1:1").SpecialCells(11).Column
    ; too lazy to look up how to convert back to alpha in VBA
    rng := "A1:" . get_excel_col(lastcol) . lastrow
    return oWorkbook.Sheets(sheet).Range(rng).Value
}
;///////////////////////////////////////////////////////////////////////////////
; Name:         make_keys
; Description:  Create a key array based on the passed format.
; Parameters:   frmt: header   == 1
;                     projects == 2
;               array: array from which to extract keys for key array
; Called by:    format_tlg
; Returns:      keyarray: array object containing keys with values of their
;                         own original index.
make_keys(frmt, array) {
    keyarray := {}
    loop % array.MaxIndex(frmt) {
        if (frmt == 1) ; projects
            key := array[A_Index, 2]
        else if (frmt == 2) ; headers
            key := array[1, A_Index]
        else msgbox,, TLG Ross - Error, %frmt% is not a valid frmt option
        keyarray.Insert(key, A_Index)
    }
    return keyarray
}
;///////////////////////////////////////////////////////////////////////////////
; Name:         format_tlg
; Description:  This function translates the tlg and description arrays into
;               usable TLG formats. Returns final TLG string to be sent to
;               calendar.
; Parameters:   tlgarr: formatted array
;               descrip:
;               xlarr:
;               xldescarr:
; Called by:    format_tlg
; Returns:      keyarray: array object containing keys with values of their
;                         own original index.
format_tlg(tlgarr, descrip, xlarr, xldescarr, defcol:=0, lastdefcol:="tr") {
    projects := make_keys(1, xlarr), headers := make_keys(2, xlarr)
    defrow:="default", row := projects[defrow], col:= headers[defcol], 
    bill := "", formatteddesc := ""
    for index, value in tlgarr {
        if headers.haskey(value) {
            col := headers[value] ; if val in header, set column number
            formatteddesc .= xldescarr[2, col] . " "
        }
        else if projects.haskey(value) {
            row := projects[value]
            formatteddesc .= xlarr[row, headers["Project"]] . " "
        }
        else if (value == "nb")
            bill := 22
        else if (value == "ed")
            bill := 7
        else return
    }

    prj := xlarr[row, headers["ID"]], tlg := xlarr[row, col]
    if (!tlg && col <= headers[lastdefcol])
        tlg := xlarr[projects[defrow], col]
    else if (col == headers[defcol] && !instr(formatteddesc, xldescarr[2, col]))
        formatteddesc .= xldescarr[2, col] . " "
    else if !tlg
        return
    return tlg . "/" . prj . "////" . bill . "," . formatteddesc
}

tlg_wrapper() {
    format_inputs(tlgarr, descrip)
    if (tlgarr == -1)
        return
    else {
        formattedtlg := format_tlg(tlgarr
                                 , descrip
                                 , __all__maintable
                                 , __all__desctable)
        return formattedtlg
    }
}
;///////////////////////////////////////////////////////////////////////////////
; Copyright © 2018 Ross F. Calimlim - LIC: GNU GPLv2
;///////////////////////////////////////////////////////////////////////////////