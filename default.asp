<!--#include file="includes/functions.asp"-->
<% pageTitle = strGreeting %>
<!--#include file="includes/header.asp"-->
<%
    strAccountNumber = Request.Form("txtAccountNumber")
    strTransactionType = Request.Form("ddlTransactionType")
    btnSubmit = Request.Form("Submit")

    If btnSubmit <> "" Then

        'submit
        Dim count
        count = Len(strAccountNumber)

        If count > 0 Then
            'add to session
            Session("AccountNumber") = strAccountNumber
            Session("TransactionType") = strTransactionType

            'validation
            If Not AccountValid(strAccountNumber) Then
                strErrorMessage = "<span style='color: #ff0000;'>Validation: Invalid Account Number, needs to have 6 or more digits.</span>"
            Else
                If DigitsOnly(count, strAccountNumber) Then
                    Response.Redirect("complete.asp")
                Else
                    strErrorMessage = "<span style='color: #ff0000;'>Validation: Invalid Account Number, digits only.</span>"
                End If
            End If

        Else
            strErrorMessage = "<span style='color: #ff0000;'>Validation: Account Number Required.</span>"
        End If

    End If

    Function AccountValid(strAccountNumber)
        If Len(strAccountNumber) >= 6 Then
            AccountValid = True
        Else
            AccountValid = False
        End If
    End Function

    Function DigitsOnly(count, strAccountNumber)
        Dim character
        For i = 1 To count
            DigitsOnly = false
            character = Mid(strAccountNumber, i, 1)
            If Not IsNumeric(character) Then Exit For
                DigitsOnly = true

        Next
    End Function
%>

<div class="content">
    <p>This page was last refreshed on <%= Now() %>.</p>

    <p><% getMessage() %></p>

    <form method="post" action="default.asp">
        Account Number: <input type="Text" name="txtAccountNumber" />

        <select name="ddlTransactionType">
            <option value="Deposit">Deposit</option>
            <option value="Withdrawl">Withdrawl</option>
            <option value="Transfer">Transfer</option>
        </select>

        <input name="Submit" type="Submit" value="Process Account" />
        <br />
        <p><%= strErrorMessage %></p>
    </form>
</div>

<!--#include file="includes/footer.asp"-->