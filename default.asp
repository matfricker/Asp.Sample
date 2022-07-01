<!--#include file="includes/functions.asp"-->
<!DOCTYPE html>
<html>
    <head>
        <title>Classic ASP - <%= strGreeting %></title>
        <link type="text/css" rel="stylesheet" href="css/styles.css" />
    </head>
    <body>
        <div class="content">
            <h3>Classic ASP - <%= strGreeting %></h3>
            <p>This page was last refreshed on <%= Now() %>.</p>
            <p><% getMessage() %></p>
            <p><a href="default.asp">Home</a></p>
        </div>

        <%
            strAccountNumber = Request.Form("txtAccountNumber")
            strTransactionType = Request.Form("ddlTransactionType")
            btnSubmit = Request.Form("Submit")

            'add to session
            Session("strAccountNumber") = strAccountNumber
            Session("strTransactionType") = strTransactionType

            If btnSubmit <> "" Then

                'SUBMIT
                Dim count
                Dim character
                Dim complete

                count = Len(strAccountNumber)
                'Response.Write(count)

                For i = 1 To count
                    character = Mid(strAccountNumber, i, 1)

                    If i = 1 Then
                        Response.Write("*")
                    Else
                        Response.Write("|")
                    End If

                    Response.Write(character)

                    If i = count Then
                        Response.Write("*")
                    End If

                Next

                If Not AccountValid(strAccountNumber) Then
                    strErrorMessage = "<span style='color: #ff0000;'>Server-Side Validation: Invalid Account Number, needs to have 6 or more digits.</span>"
                Else
                    Server.Transfer("complete.asp")
                End If

            End If

            Function AccountValid(strAccountNumber)
                If Len(strAccountNumber) >= 6 Then
                    'For Each item In strAccountNumber.Char
                    AccountValid = True
                Else
                    AccountValid = False
                End If
            End Function
        %>

        <div class="content">
            <form method="post" action="default.asp">
                Account Number: <input type="Text" name="txtAccountNumber" />

                <select name="ddlTransactionType">
                    <option value="Deposit">Deposit</option>
                    <option value="Withdrawl">Withdrawl</option>
                    <option value="Transfer">Transfer</option>
                </select>

                <input name="Submit" type="Submit" value="Submit" />
                <br />
                <p><%= strErrorMessage %></p>
            </form>
        </div>
    </body>
</html>