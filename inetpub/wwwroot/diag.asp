<%@ EnableSessionState=False %>
<%
Dim conn, rs, connStr
connStr = "Provider=SQLOLEDB;Data Source=DESKTOP-S29QIQ5\SQLEXPRESS;Initial Catalog=v4_stage;user ID=SA;password=algoraSaturn2203@;"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connStr

Response.Write "<h2>Lunia Server Diagnostics</h2>"

Response.Write "<h3>StageServers (servidores registrados):</h3>"
Set rs = conn.Execute("SELECT * FROM StageServers")
If rs.EOF Then
    Response.Write "<p style='color:red'>VAZIO - Nenhum servidor registrado!</p>"
Else
    Response.Write "<table border=1><tr>"
    Dim f
    For Each f In rs.Fields
        Response.Write "<th>" & f.Name & "</th>"
    Next
    Response.Write "</tr>"
    Do Until rs.EOF
        Response.Write "<tr>"
        For Each f In rs.Fields
            Response.Write "<td>" & f.Value & "</td>"
        Next
        Response.Write "</tr>"
        rs.MoveNext
    Loop
    Response.Write "</table>"
End If
rs.Close

Response.Write "<h3>SquareStages:</h3>"
Set rs = conn.Execute("SELECT * FROM SquareStages")
If rs.EOF Then
    Response.Write "<p style='color:red'>VAZIO - Nenhuma square registrada!</p>"
Else
    Response.Write "<table border=1><tr>"
    For Each f In rs.Fields
        Response.Write "<th>" & f.Name & "</th>"
    Next
    Response.Write "</tr>"
    Do Until rs.EOF
        Response.Write "<tr>"
        For Each f In rs.Fields
            Response.Write "<td>" & f.Value & "</td>"
        Next
        Response.Write "</tr>"
        rs.MoveNext
    Loop
    Response.Write "</table>"
End If
rs.Close

Response.Write "<h3>Connections:</h3>"
Set rs = conn.Execute("SELECT COUNT(*) as total FROM Connections")
Response.Write "<p>Total de conexoes: " & rs(0) & "</p>"
rs.Close

Response.Write "<h3>Databases disponiveis:</h3>"
Set rs = conn.Execute("SELECT name FROM sys.databases WHERE name IN ('v3_character','v4_stage','v3_guild','v3_gate','d-shop','crmdb') ORDER BY name")
Do Until rs.EOF
    Response.Write "<p style='color:green'>" & rs(0) & " OK</p>"
    rs.MoveNext
Loop
rs.Close

conn.Close
Set conn = Nothing
%>
