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
; This script translates shorthand TLG information entered by a user into
; properly formatted time logging strings. This script, in general, defaults
; to silent errors and returns nothing due to usability concerns when keying
; information in quick succession.
;
; Update: 2018 November 3

;//////////////////////////////////////////////////////////////////////////////
; DEFINE GLOBALS
;//////////////////////////////////////////////////////////////////////////////

; make a safe array from passed excel matrix and sheet name
global __all__maintable := make_safe_arr("D:\Documents\matrix.xlsx", "Main")

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
; Parameters:   none
; Called by:    format_inputs
; Returns:      string (if user enters information)
;               -1 (otherwise)
get_input() {
    msg = org + tlp + bill, desc
    inputbox, str, TLG Ross, %msg%,, 150, 130 ; inputbox size 200x150
    if (errorlevel != 0 || str == "") { ; return -1 if anything but user entry
        return -1
    }
    else return str ; otherwise return the string input
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         str_to_arr
; Description:  Converts string to array using passed delimeter. Also omits
;               passed characters.
; Parameters:   str:   string to create array from
;               delim: delimiter string (defaults to nothing so every char is
;                      parsed)
;               omit:  characters to exclude from strings (defaults to nothing)
; Called by:    format_inputs
; Returns:      arr: if string exists
;               -1: if string is empty or an error (-1)
str_to_arr(str, delim:="", omit:="") {
    if (!str || str == -1) {
        return -1
    }
    else {
        return arr := strsplit(str, delim, omit)
    }
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         arr_to_str
; Description:  Flattens array into human readable string. Good for testing,
;               not actually called for tlg process.
; Parameters:   arr: array to convert to string
; Called by:    format_inputs
; Returns:      string
arr_to_str(arr) {
    string := "{"
    for key, value in arr {
        if(A_index != 1)
            string .= ","
        if key is number
            string .= key ":"
        else if(isobject(key))
            string .= arr_to_str(key) ":"
        else {
            string .= key . ":"
        }
        if value is number
            string .= value
        else if (isobject(value))
            string .= arr_to_str(value)
        else
            string .= value
    }
    return string . "}"
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
format_inputs(byref tlg_arr, byref des_str) {
    user_inp := get_input()
    if (user_inp == -1 or user_inp == "")
        tlg_arr := des_str := -1
    else {
        has_comma := instr(user_inp, ",")
        if (!has_comma)
            comma_pos := strlen(user_inp)
        else comma_pos := has_comma
        tlg := substr(user_inp, 1, comma_pos)
        , des := substr(user_inp, comma_pos + 1)
        , tlg_arr := str_to_arr(tlg, " ", ",")
        , des_str := trim(des)
    }
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         num_to_alpha
; Description:  Takes an integer and returns its alphabetic equivalent. Errors
;               if parameter is a non-integer or outside of range 1-26.
; Parameters:   int: integer to convert to alpha character
; Called by:    excel_encode
; Returns:      single alphabetic character (good input)
;               error message (bad input)
num_to_alpha(int) {
    alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if int is not integer
        return "Non-integer input"
    else if (int < 0 || !int || int > 27)
        return "Integer out of alphabetic bounds"
    else return substr(alphabet, int, 1)
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         get_excel_col
; Description:  I cannot for the life of me figure out how VBA works, so I had
;               to write a function that converts the numeric column returned
;               from a SpecialCell lookup into familiar alphabetic excel
;               notation. This is a recursive function.
; Parameters:   column_num: Excel numeric column ID
;               divisor:    modulo divisor (should always be 26 but whatever)
; Called by:    excel_encode (recursively)
;               make_safe_arr
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
        , column_num := floor(column_num/26)
        return % excel_encode(column_num) . encode_num(remainder)
    }
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         make_safe_arr
; Description:  Gets an excel workbook from passed file path and returns a safe
;               array object.
; Parameters:   sheet:     sheet name, defaults to 1
;               file_path: path to excel matrix
; Called by:    __all__maintable (global)
; Returns:      array object
make_table(sheet, file_path := "C:\Users\Ross\Desktop\matrix.xlsx") {
    oWorkbook := comobjget(file_path)
    , lastrow := oWorkbook.Sheets(sheet).Range("A:A").SpecialCells(11).Row
    , lastcol := oWorkbook.Sheets(sheet).Range("1:1").SpecialCells(11).Column
    ; too lazy to look up how to convert back to alpha in VBA
    rng := "A1:" . get_excel_col(lastcol) . lastrow
    return oWorkbook.Sheets(sheet).Range(rng).Value
    ; return oWorkbook.Sheets("Main").Range("A1:O11").Value
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         make_keys
; Description:  Create a key array based on the passed format.
; Parameters:   frmt: header   == 1
;                     projects == 2
;               array: array from which to extract keys for key array
; Called by:    format_tlg
; Returns:      key_array: array object containing keys with values of their
;                         own original index
make_key_arr(array, frmt) {
    key_array := {}
    if (frmt == "row") {
        loop % array.maxindex(2) {
            key := array[1, a_index], val := array[2, a_index]
            , arr_val := {"index": a_index, "description": val}
            key_array.insert(key, {"index": a_index, "description": val})
        }
        return key_array
    }
    else if (frmt == "col") {
        loop % array.maxindex(1) {
            key := array[a_index, 2], val := array[a_index, 1]
            key_array.insert(key, {"index": a_index, "description": val})
        }
        return key_array
    }
    else
        msgbox, "Format must be row or col"
        return
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         format_tlg
; Description:  This function translates the tlg and description arrays into
;               usable TLG formats. Returns final TLG string to be sent to
;               calendar.
; Parameters:   tlgarr:     formatted array
;               descrip:
;               xlarr:
;               xldescarr:
; Called by:    format_tlg
; Returns:      key_array: array object containing keys with values of their
;                         own original index.
format_tlg(safe_arr, tlg_arr, des_str, def_row, def_col, last_def_col) {
    func := "format_tlg"
    
    ; Initial iterative formatting--bulk of the final output
    proj_arr := make_key_arr(safe_arr, "col")
    , head_arr := make_key_arr(safe_arr, "row")
    , row_num := proj_arr[def_row]["index"]
    , col_num := head_arr[def_col]["index"]
    , tlg_desc := des_str, tlg_bill := "", tlg_bill_des := ""
    , bill_arr := {"nb": {"index": 22, "description": "non-bill"}
                 , "ed": {"index": 7, "description": "education"}}
    
    for key, value in tlg_arr {
        if % head_arr.haskey(value) {
            col_num := head_arr[value]["index"]
            if (!des_str)
                tlg_desc .= head_arr[value]["description"] . " "
        }
        else if % proj_arr.haskey(value) {
            row_num := proj_arr[value]["index"]
            if (!des_str)
                tlg_desc .= proj_arr[value]["description"] . " "
        }
        else if % bill_arr.haskey(value) {
            tlg_bill := bill_arr[value]["index"]
            tlg_desc .= " " . bill_arr[value]["description"]
        }
        else return
    }
    
    ; Define final tlg values (minor formatting tweaks and edge cases)
    def_row_num := proj_arr[def_row]["index"]
    , def_col_num := head_arr[last_def_col]["index"]
    , prj := safe_arr[row_num, head_arr["ID"]["index"]]
    if (col_num  <= def_col_num) { ; assign def tlp ID if col within def range
        tlp := safe_arr[def_row_num, col_num]
    }
    else {
        tlp := safe_arr[row_num, col_num]
    }
    return % tlp . "/" . prj . "////" . tlg_bill . "," . tlg_desc
}
;//////////////////////////////////////////////////////////////////////////////
; Name:         tlg_wrapper
; Description:  Wraps all the tlg functions together.
; Parameters:   safe_arr:     array from excel file
;               def_row:      default row mnemonic
;               def_col:      default column abbreviation
;               last_def_col: last column abbreviation to be considered for
;                             for default assignment functionality
; Called by:    this script
; Returns:      final_tlg:    formatted tlg string
tlg_wrapper(safe_arr, def_row, def_col, last_def_col) {
    func := "tlg_wrapper"
    format_inputs(tlg_arr, des_str)
    if (tlg_arr == -1)
        return
    else {
        return final_tlg := format_tlg(safe_arr
                              , tlg_arr
                              , des_str
                              , def_row
                              , def_col
                              , last_def_col)
    }
}
;//////////////////////////////////////////////////////////////////////////////
; Copyright © 2018 Ross F. Calimlim - LIC: GNU GPLv2
;//////////////////////////////////////////////////////////////////////////////