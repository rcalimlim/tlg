;//////////////////////////////////////////////////////////////////////////////
; SCRIPT SETTINGS
;//////////////////////////////////////////////////////////////////////////////
#SingleInstance force
#NoEnv ; Recommended for performance and compatibility with future 
       ; AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior 
               ; speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.

;//////////////////////////////////////////////////////////////////////////////
; TLG ROSS v3.4
;//////////////////////////////////////////////////////////////////////////////
; This script translates shorthand TLG information entered by the user into
; proper delorean codes. This script, in general, defaults to silent errors
; and returns nothing due to usability concerns when keying information in 
; quick succession.
;
; Update: 2018 October 30

;//////////////////////////////////////////////////////////////////////////////
; DEFINE GLOBALS
;//////////////////////////////////////////////////////////////////////////////
global __all__maintable := make_safe_arr("D:\Documents\matrix.xlsx")

;//////////////////////////////////////////////////////////////////////////////
; DEFINE HOTKEYS 
;//////////////////////////////////////////////////////////////////////////////
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

;//////////////////////////////////////////////////////////////////////////////
; DEFINE FUNCTIONS
;//////////////////////////////////////////////////////////////////////////////
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
;//////////////////////////////////////////////////////////////////////////////
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
;//////////////////////////////////////////////////////////////////////////////
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
;//////////////////////////////////////////////////////////////////////////////
; Name:         encode_num
; Description:  Takes an integer and returns its alphabetic equivalent. Errors
;               passed value is not an integer or not within 1-26.
; Parameters:   int: integer to convert to alpha character
; Called by:    excel_encode
; Returns:      alphabetic character (good input)
;               error message (bad input)
encode_num(int) {
    alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if int is not integer
        return "Non-integer input"
    else if (int < 0 || !int || int > 27)
        return "Integer out of alphabetic bounds"
    else return substr(alphabet, int, 1)
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         excel_encode
; Description:  I cannot for the life of me figure out how VBA works, so I had
;               to write a function that converts the numeric column returned
;               from a SpecialCell lookup into familiar alphabetic excel
;               notation. 
;               This is a recursive function.
; Parameters:   column_num: Excel numeric column ID
;               divisor: modulo divisor (should always be 26 but whatever)
; Called by:    excel_encode (recursively)
;               make_safe_arr
; Returns:      alphabetic translation of col ID (good input)
;               error message (bad input)
excel_encode(column_num) {
    errormsg := "Parameters must be positive integers"
    if column_num is not integer
        return % errormsg
    else if (column_num <= 0)
        return % errormsg
    else if (column_num <= 26)
        return % encode_num(column_num)
    else {
        remainder := mod(column_num, 26)
        column_num := floor(column_num/26)
        return % excel_encode(column_num) . encode_num(remainder)
    }
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         make_safe_arr
; Description:  Gets an excel workbook from passed file path and returns a safe
;               array object.
; Parameters:   sheet: sheet name, defaults to 1
;               file_path: path to excel matrix
; Called by:    __all__maintable (global)
;               __all__desctable (global)
; Returns:      array object
make_safe_arr(file_path, sheet:=1) {
    oWorkbook := comobjget(file_path)
    ; VBA crap probably
    lastrow := oWorkbook.Sheets(sheet).Range("A:A").SpecialCells(11).Row
    lastcol := oWorkbook.Sheets(sheet).Range("1:1").SpecialCells(11).Column
    ; too lazy to look up how to convert back to alpha in VBA
    rng := "A1:" . excel_encode(lastcol) . lastrow
    return oWorkbook.Sheets(sheet).Range(rng).Value
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         make_key_arr
; Description:  Create a key array based on the passed format.
; Parameters:   array: array from which to extract keys for key array
;               frmt:  col = assign values from column
;                      row = assign values from row
; Called by:    format_tlg
; Returns:      keyarray: array object containing keys with values of their
;                         own original index
make_key_arr(array, frmt) {
    keyarray := {}
    if (frmt == "row") {
        loop % array.maxindex(1) {
            key := array[1, a_index]
            val := array[2, a_index]
            keyarray.insert(key: {"index": a_index, "description": val})
        }
        return keyarray
    }
    else if (frmt == "col") {
        loop % array.maxindex(2) {
            key := array[a_index, 2]
            val := array[a_index, 1]
            keyarray.insert(key: {"index": a_index, "description": val})
        }
        return keyarray
    }
    else return -1
    }
}
;//////////////////////////////////////////////////////////////////////////////
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
    projects := make_key_arr(xlarr, 1), headers := make_key_arr(xlarr, 2)
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
    else if (col == headers[defcol] && !instr(formatteddesc
                                            , xldescarr[2, col]))
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
;//////////////////////////////////////////////////////////////////////////////
; Copyright © 2018 Ross F. Calimlim - LIC: GNU GPLv3
;//////////////////////////////////////////////////////////////////////////////