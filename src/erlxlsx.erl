-module(erlxlsx).
-export([read/1, write/3]).
-include_lib("xmerl/include/xmerl.hrl").
-define(HEAD, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>").
-define(XMLNS, "http://schemas.openxmlformats.org/spreadsheetml/2006/main").
-define(XMLNSR, "http://schemas.openxmlformats.org/officeDocument/2006/relationships").

read(File) ->
	zip:unzip(File),
	Strings = readSharedStrings("xl/sharedStrings.xml"),
	readSheets("xl/worksheets/sheet1.xml", Strings).

write(FileName, Columns, Data) ->
	Strings = findSharedStrings(Columns ++ lists:concat(Data), []),
	writeToFile("../priv/template/xl/sharedStrings.xml", createSharedStringsXML(Strings)),
	writeToFile("../priv/template/xl/worksheets/sheet1.xml", createSheet([Columns] ++ Data, Strings)),
	zipTemplateFile(FileName).

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
	Sst = {sst, [{xmlns,?XMLNS}, {count,Count}, {uniqueCount,Count}], SiList},
	xmerl:export_simple([Sst], xmerl_xml).

createSheet(Data, Strings) ->
	Rows = createRowXML(Data, Strings, [], 0),
	WorkSheet = {worksheet, [{xmlns, ?XMLNS}, {'xmlns:r', ?XMLNSR}], [{sheetData, [], Rows}]},
	xmerl:export_simple([WorkSheet], xmerl_xml).

createRowXML([], _Strings, List, _RowNum) ->
	List;

createRowXML([H|T], Strings, List, RowNum) ->
	NewRowNum  = RowNum + 1,
	RownNumStr = integer_to_list(NewRowNum),
	Columns    = createColumnXML(H, Strings, [], RownNumStr, 1),
	Row = {row, [{r, RownNumStr}, {customFormat,"false"}, {ht,"13.85"}, {hidden,"false"}, 
				{customHeight,"false"}, {outlineLevel,"0"}, {collapsed,"false"}], Columns},
	createRowXML(T, Strings, List ++ [Row], NewRowNum).

createColumnXML([], _Strings, List, _RowNum, _ColumnCount) ->
	List;

createColumnXML([H|T], Strings, List, RowNum, ColumnCount) ->
	CoulmnNum = lists:flatten(io_lib:format("~c", [64 + ColumnCount])) ++ RowNum,
	Column = case is_integer(H) of
				true  -> {c, [{r,CoulmnNum}, {s,0}, {t, "n"}], [{v, [], [integer_to_list(H)]}]};
				false -> SharedStringId = string:str(Strings, [H]),
						 {c, [{r,CoulmnNum}, {s,0}, {t, "s"}], [{v, [], [integer_to_list(SharedStringId)]}]}
			 end,
	createColumnXML(T, Strings, List ++ [Column], RowNum, ColumnCount + 1).


createSharedStringSi([], List) ->
	List;

createSharedStringSi([H|T], List) ->
	Si = {si, [], [{t, [], [H]}]},
	createSharedStringSi(T, List ++ [Si]).

writeToFile(File, String) ->
	file:write_file(File, String). 

zipTemplateFile(FileName) ->
	zip:zip(FileName, ["_rels/.rels","docProps/app.xml","docProps/core.xml",
     "xl/_rels/workbook.xml.rels","xl/sharedStrings.xml",
     "xl/worksheets/sheet1.xml","xl/styles.xml",
     "xl/workbook.xml","[Content_Types].xml"], [{cwd, "../priv/template/"}, verbose]),
	io:format("~s has been created~n", [FileName]).

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
