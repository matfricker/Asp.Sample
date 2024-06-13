<!--#include file="includes/functions.asp"-->
<% pageTitle = strGreeting %>
<!--#include file="includes/header.asp"-->
<%
    strAccountNumber = Request.Form("txtAccountNumber")
    strTransactionType = Request.Form("ddlTransactionType")

    If Not IsEmpty(Request.Form) Then

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
            <option value="1">Credit</option>
            <option value="2">Debit</option>
        </select>

        <input name="Submit" type="Submit" value="Process Account" />
        <br />
        <p><%= strErrorMessage %></p>
    </form>
</div>

<div class="content">

    <form method="post" action="generate-xml.asp">
        <label>Generate XML File:</label>
        <input name="Submit" type="Submit" value="Generate" />
    </form>
</div>

<!--#include file="includes/footer.asp"-->