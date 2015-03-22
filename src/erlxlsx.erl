-module(erlxlsx).
-export([read/1, write/3]).
-include_lib("xmerl/include/xmerl.hrl").
-define(HEAD, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>").

read(File) ->
	zip:unzip(File),
	Strings = readSharedStrings("xl/sharedStrings.xml"),
	readSheets("xl/worksheets/sheet1.xml", Strings).

write(FileName, Columns, Data) ->
	io:format("~p~n", [lists:concat(Columns ++ Data)]),
	Strings = findSharedStrings(Columns ++ lists:concat(Data), []),
	createSharedStringsXML(Strings).

%write functions
findSharedStrings([], Strings) ->
	Strings;

findSharedStrings([H|T], Strings) ->
	case (is_integer(H) or lists:member(H, Strings)) of
		true  -> no_op,
				 findSharedStrings(T, Strings);
		false -> findSharedStrings(T, Strings ++ [H])
	end.

createSharedStringsXML(Strings) ->
	Count = integer_to_list(length(Strings)),
	SiList = createSharedStringSi(Strings, []), 
	Sst = {sst, [{xmlns,"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}, {count,Count}, {uniqueCount,Count}], SiList},
	?HEAD ++ lists:flatten(xmerl:export_simple_element(Sst, xmerl_xml)).

createSharedStringSi([], List) ->
	List;

createSharedStringSi([H|T], List) ->
	Si = {si, [], [{t, [], [H]}]},
	createSharedStringSi(T, List ++ [Si]).
 

%read functions
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
