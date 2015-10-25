<%
Function QueryToJSON(dbc, sqlJSONsys)
        Dim rsJSONsys, jsa
        Set rsJSONsys = dbc.Execute(sqlJSONsys)
        Set jsa = jsArray()
        While Not (rsJSONsys.EOF Or rsJSONsys.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rsJSONsys.Fields
                        jsa(Null)(col.Name) = col.Value
                Next
        rsJSONsys.MoveNext
        Wend
        Set QueryToJSON = jsa
End Function
%>