# erlxlsx
Erlang library to encode and decode data into xlsx file

Currently only decoding is complete, which only converts the first sheet in the xlsx file

###Usage : 

    erlxlsx:read(XlsxFilePath).

Sample Output 

>erlxlsx:read("../priv/test.xlsx").
[["id","value","attribute"],
 [1,"one","I"],
 [2,"two","ii"],
 [3,"three","iii"],
 [4,"four","iv"],
 [5,"five","v"]]

would return data as list of lists.

    erlxlsx:write(FileName, Columns, Data).

would create xlsx file in the same folder.

Example

>erlxlsx:write("test4.xlsx", ["id", "value", "attribute"], [[1, "one", "i"],[2,"two","ii"],[3,"three","iii"],[4,"four","iv"],[5,"five","v"], [6, "six", "vi"]]).

###TODO:

* Multiple Sheet compatibility
* Multiple Data type
* Column Names after Z





