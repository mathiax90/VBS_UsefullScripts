call main

sub main	
	Set objRegExp = CreateObject("VBScript.RegExp")
	objRegExp.Pattern = "^OUT_\d{4}_\d{2}_\d{2}.txt$"
	msgbox objRegExp.test("OUT_2020_12_31.txt")	
end sub