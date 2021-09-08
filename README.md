# Excel_VBA_Lib
Library of functions for Excel VBA

I hope this will be useful. Especially the sitring library since strings are so critical 
for working with formulas and references.

Below is an index of the functions:

 #############################################################################
 Key_Val library

 a key-value structure stored in a string
 e.g. "one:1,two:2,three:3"
 assumed no duplicate keys


 keyval_put(keyval, key, val)		create
 keyval_get(key, data) 			read 
 keyval_update(data, key, new_val)	update
 keyval_delete(data, key)			delete
 keyval_key_exists(data, key)		returns whether key is in data

 #############################################################################
 Date Library

 monthShort(n) 			given a month as an int, returns short form
 monthLong(n)  			given a month as an int, returns long form
 monthStringtoInt(string)	given a month in long form, returns number
 lastDayofMonth(n) 		given a month as a number, returns last day of the month

 #############################################################################

 function library for ranges
 dependencies:
   string

 visibleCellsCountNotZero(range)	returns whether the count of visible cells within a range is NOT zero
 sumRange(string) 			given an address, returns sum of values within that range
 firstCell(range)			returns address of the FIRST cell within a range
 lastcell(range)			returns address of LAST cell within a range
 nRow(range)				returns count of rows in provided range
 nCols(range)				returns count of columns in a provided range
 contiguousRange()			from selection, returns the address of all the contiguous range up, down, left, right
 nextCol(range, offset)	returns the address of a range of the next sub-column in provided range given offset

 #############################################################################

function library for cell references for excel vba
 dependency libraries:
   string

 isReference(formula) 			returns wheteher or not a tab is in referenced in the formula
 isFormula(string)    			returns whether or not a string contains a formula
 hasParentRef(list, delimiter)	given a list of cell addresses, checks if tab is referenced or not
 issequential(addr1, addr2) 	returns whether two addresses are next to each other
 getRef(formula) 				gets a formula and returns workbook reference if it contains
 removeReferences() 			remove references to other workbooks in selection
 convert_range_to_list(string)	converts a range demarked by ":" into a list of addresses
 highlightDereferencedCells()	Dereference cell references in selection until a cell with a value is found. 
					   				 Then highlight it.
 reflistSuccincttoVerbose(s) 	Gets a cell formula and expands all ranges (A1:A3,A5:A6) of referenced cells
                                    into one monster list(A1,A2,A3,A5,A6...etc.)
 refListVerboseToSuccinct(list, delim)
								formats the input (a list of references) as output: =sum(sheet!a1:a3,sheet!a5)
 refcnt(string) 				returns count of references in formula
 getNthRef(n, formula) 			returns the nth reference in the formula
 convert_sum_to_ref_list(string)	
								another succinct to verbose, but this one for sum formulas
 create_cell_list(range)		creates a list (contained in a string) from the range of cells provided
 create_cell_list_if(criteria, range, col) 
								returns a list of formatted cell addresses that match the criteria function
                               	provided, assumed references contain tab names
 paste_remove_references()		pastes and removes workbook references
 rmvRefs(f)						removes workbook references from formula
 offset_cell_row_reference()	offset the cell reference row
 vlookup_refs_only(val, range, offset) 
 								like vlookup in Excel, but returns reference instead of value on first match

 #############################################################################

 functions pertaining to rows and columns

 has_row(search_text) 		  returns whether or not search_text can be found in given sheet, searches by row
 lastcol(row) 				  returns last column in row
 lastrow(column) 			  returns last row in column
 colToLetter(c)			  	  converts a column as a number to a string
 letterToCol(c)			  	  converts column as a letter to a number
 getMaxRow()				  returns the max row on the sheet
 firstRow(c, offset)		  returns the first row that has a value in the column. 
                                  Offset can be used to ignore headers
 findColAfter(sheet, search_text, after) 
 							  returns the first column containing search_text after range
 getRowStart(value, range) 	  returns first row that equals value in the range
 getRowEnd(value, range)      returns the last row that equals value in the range
 findFirstCol(sheet, search_text) 
 							  returns first column containing search_text on the sheet
 findFirstRow(sheet, search_text) 
 							  returns first row containing search_text on the sheet
 isFound(search_text)		  returns whether search_text exists on the ActiveSheet
 findCumulativeTotalRow(range, amt, write_total) returns total row; 
       						      write_total = true writes out cumulative totals next to range

 #############################################################################

 String Library for excel vba

 isalpha(char)  				receives a character and returns whether or not it is a letter
 is_integer(char) 				receives a character and returns wheter or not it is an integer
 doesNotContain(string, chars) 	returns if string contains the supplied chars
 has_alpha(string)        		returns whether the string has ANY letters
 find(find_text, within_text, start_num) 	returns index of find_text in within_text starting from the start_num
 word_count(string, word) 		returns the count of word in string
 remove_word(string, word)		returns a string omitting all instances of word found in received string
 replace_word(string, word, replacement)
 								returns a string where instances of word is replaced by replacement
 word_count_range(rng, word) 	for a range, counts all instances of word in cell contents
 not_numbers(string)			given a string, returns all but numbers
 numbers_only(string)			given a string, returns only numbers
 remove_char(string, char)		given a string, returns all but char
 nextAlpha(char)				given a character, returns the next character in the alphabet
 char_diff(char1, char2)		given two characters, returns the count between them in the alphabet
 alphaOnly(string)				given a string, returns only the letters
 int_to_char(integer)			converts an integer to a character
 char_to_int(char)				converts a character to an integer
 upper(string)					returns upcased version of string
 remove_all_between(string, start, stop) 	
 								returns a string with all removed between the start and stop chars
 onlyBeforeDelimiter(string, char) 		
 								returns a string with only characters before a character is reached
 rmvFirstChar(string, char) 	returns a string where the first instance of char is removed
 integer_breakout(string, n) 	returns the nth integer found int he string received
 n_integers(s As String) 		returns the number of integers found in the received string, NOT # of digits
 printAfter(string, index) 		returns string where all is removed except characters after index
 printAfterNth(find_text, within_text, n) 	
 								return string after the nth find_text in within_text
 findNth(find_text, within_text, n) 		
 								returns index of nth find_Text within provided text
 slice(string, start, stop) 	returns a string starting with start and ending with stop
 nWords(string) 				returns number of words containing 3 or more letters in string
 nWhtieSpace(string) 			returns number of spaces in string
 getNthWord(string, n) 			returns nth word of a string after (assuming word contains 3 letters or more)
 remove_duplicates(string_list, delimiter)	
 								removes duplicates from a string list where members are separated by delimiter

 #############################################################################

 function library for tabs


 doesTabExist(name) 		returns whether tab exists
 find_tab_index(keyword) 	returns index of tab that matches keyword

 #############################################################################

 workbook library

 activatewb(wbName) 		activates workbook by name and returns whether wb was found
 openwb(path, file) 		opens workbook given a path and file and returns success or failure
 openwb_path_only(path) 	opens workbook given a full file path and returns success or failure
