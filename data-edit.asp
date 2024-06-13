<!--#include file="includes/functions.asp"-->
<% pageTitle = "SQLite Data - Edit" %>
<!--#include file="includes/header.asp"-->
<%
    Dim accountId
    accountId = Request.QueryString("id")

    ' connection and recordset variables
    Dim conn, strConn
    Dim strInsertTransaction, strTransactionSequence, strAccountByAccountId
    Dim rsSequence, rsAccounts
    Dim fld, Err

    ' open connection
    Set conn = Server.CreateObject("ADODB.Connection")
    strConn = "DRIVER=SQLite3 ODBC Driver;Database=" & Server.MapPath("data\data.sqlite3") & ";LongNames=0;Timeout=1000;NoTXN=0;SyncPragma=NORMAL;StepAPI=0;"
    conn.Open strConn

    'If this is first time page is open, Form collection
    'will be empty when data is entered. run AddNew method
    If Not IsEmpty(Request.Form) Then
        If Not Request.Form("AccountId") = "" Then

            Dim nextId, id, transactionType, transactionAmount, transactionDate

            id = Request.Form("AccountId")
            transactionType = Request.Form("TransactionType")
            transactionAmount = Request.Form("Amount")

            Dim today, formattedDate
            today = Now
            formattedDate = Year(today) & "-" & addLeadingZero(Month(today)) & "-" & addLeadingZero(Day(today)) & " " & addLeadingZero(Hour(today)) & ":" & addLeadingZero(Minute(today))
            transactionDate = formattedDate ' yyyy-mm-dd hh:mm

            strTransactionSequence = "SELECT seq FROM sqlite_sequence WHERE name = 'Transactions'"
            Set rsSequence = Server.CreateObject("ADODB.recordset")
            rsSequence.Open strTransactionSequence, conn

            nextId = rsSequence("seq") + 1
            strInsertTransaction = "INSERT INTO Transactions VALUES ("& nextId &","& id &","& transactionType &","& transactionAmount &",'"& transactionDate &"')"
            conn.Execute strInsertTransaction

             ' check for errors
            If conn.Errors.Count > 0 Then
                For Each Err In conn.Errors
                    Response.Write("Error " & Err.SQLState & ": " & _
                        Err.Description & " | " & Err.NativeError)
                Next
                conn.Errors.Clear
            Else
                Response.Redirect("data.asp?id=" & accountId)
            End If
        End If
    End If

%>
<div class="content">
    <%
        If accountId <> "" Then
            strAccountByAccountId = "SELECT * FROM Accounts WHERE Id = " & accountId
            Set rsAccounts = Server.CreateObject("ADODB.recordset")
            rsAccounts.Open strAccountByAccountId, conn
            Response.Write("<p><strong></strong></p>")
        End If
    %>

    <form method="post" action="data-edit.asp">
        <h3>New Transaction:</h3>

        <input type="hidden" name="AccountId" value="<%= accountId%>" style="height: 16px; width: 82px;"/>

        <div style="margin-bottom: 8px;">
            <label style="display: inline-block; width: 150px;">Customer:</label>
            <strong><%= rsAccounts("FirstName") & " " & rsAccounts("LastName") %></strong>
        </div>

        <div style="margin-bottom: 8px;">
            <label style="display: inline-block; width: 150px;">Transaction Type:</label>
            <select title="Transaction Type" name="TransactionType" style="height: 20px; width: 90px;">
                <option value="1">Credit</option>
                <option value="2">Debit</option>
            </select>
        </div>

        <div style="margin-bottom: 8px;">
            <label style="display: inline-block; width: 150px;">Amount:</label>
            <input type="textbox" name="Amount" style="height: 16px; width: 82px;"/>
        </div>

        <div>
            <label style="display: inline-block; width: 150px;">&nbsp;</label>
            <input name="Submit" type="Submit" value="Add Transaction" />
        </div>

        <div>
            <p><%= strErrorMessage %></p>
        </div>

    </form>
</div>

<%
    ' Clean up.
    ''If rsAccounts.State = adStateOpen Then
    ''   rsAccounts.Close
    ''End If

    'If rsSequence.State = adStateOpen Then
    '   rsSequence.Close
    'End If

    If conn.State = adStateOpen Then
       conn.Close
    End If

    Set rsAccounts = Nothing
    Set rsSequence = Nothing
    Set conn = Nothing
    Set fld = Nothing
%>

<!--#include file="includes/footer.asp"-->