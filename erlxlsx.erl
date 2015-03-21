-module(erlxlsx).
-export([getData/1]).
-include_lib("xmerl/include/xmerl.hrl").

getData(File) ->
	zip:unzip(File),
	Strings = readSharedStrings("xl/sharedStrings.xml"),
	readSheets("xl/worksheets/sheet1.xml", Strings).

getXMLData(File) ->
	Data = case xmerl_scan:file(File) of
				{XMLData, []} 	 -> XMLData;
				{_XMLData, Misc} -> exit({trailing_data, Misc})
		   end,
	Data.

readSharedStrings(SharedStringsFile) ->
	StringData = getXMLData(SharedStringsFile),
	getStrings(xmerl_xpath:string("//si/t/text()", StringData), []).

getStrings([], Strings) ->
	Strings;

getStrings([H | T], Strings) ->
	getStrings(T, Strings ++ [getText(H)]).

readSheets(SheetFile, Strings) ->
	SheetData = getXMLData(SheetFile),
	Rows = xmerl_xpath:string("//worksheet/sheetData/row", SheetData),
	processRows(Rows, Strings, []).

processRows([], _Strings, Data) ->
	Data;

processRows([H|T], Strings, Data) ->
	Columns = processColumns(xmerl_xpath:string("/row/c", H), Strings, []),
	processRows(T, Strings, Data ++ Columns).

processColumns([], _Strings, Columns) ->
	[Columns];

processColumns([H|T], Strings, Columns) ->
	[Text] = xmerl_xpath:string("/c/v/text()", H),
	NewColumns = case xmerl_xpath:string("/c[@t='n']", H) of
					[] 	    -> Columns ++ [lists:nth(getText(Text, int) + 1, Strings)];
					_Number -> Columns ++ [getText(Text, int)]
				 end,
	processColumns(T, Strings, NewColumns).

getText(XMLText)  when is_record(XMLText,xmlText) ->
	XMLText#xmlText.value.

getText(XMLText, int) ->
	list_to_integer(getText(XMLText)).
