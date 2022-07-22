<!--#include file="includes/functions.asp"-->
<% pageTitle = "SQLite Data" %>
<!--#include file="includes/header.asp"-->

<div class="content">
    <p>The following is extracted from a SQLite database, using the SQLite3 ODBC Driver.</p>
    <%
        Set conn = Server.CreateObject("ADODB.Connection")

        Dim path 
        path = Server.MapPath("data\data.sqlite3")
        conn.Open "DRIVER=SQLite3 ODBC Driver;Database=" & path & ";LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;"

        Set recordSet = Server.CreateObject("ADODB.recordset")

        Dim sqlStatementCount
        sqlStatementCount = "SELECT COUNT(*) As Count FROM Accounts"

        Set recordSetCount = conn.Execute(sqlStatementCount)

        Dim recordCount
        If Not recordSetCount.EOF Then
            recordCount = recordSetCount.Fields("Count")
        Else
            recordCount = 0
        End IF

        If recordCount > 0 Then

            Dim sqlStatement
            sqlStatement = "SELECT * FROM Accounts"

            recordSet.Open sqlStatement, conn

            Response.Write("<p>Accounts: (" & recordCount & ")</p>")
            Response.Write("<ul class='account-list'>")

            Do While Not recordSet.EOF
                Response.Write("<li>" & cstr(recordSet("AccountNumber")) & " | " & recordSet("FirstName") & " " & recordSet("LastName") & "</li>")
                recordSet.MoveNext
            Loop

            Response.Write("</ul>")
        Else
            Response.Write("<p>No Records</p>")
        End If
    %>
</div>

<!--#include file="includes/footer.asp"-->