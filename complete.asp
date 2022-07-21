<!--#include file="includes/functions.asp"-->
<% pageTitle = "Complete" %>
<!--#include file="includes/header.asp"-->

<div class="content">
    <p>Process Complete</p>

    <%
        If Session("AccountNumber") <> "" Then
            Dim str 
            str1 = "<p>Account Number: AC-" + Session("AccountNumber") + "-XXXX</p>"
            Response.Write(str1)

            str2 = "<p>Transaction Type: " + Session("TransactionType") + "</p>"
            Response.Write(str2)
        End If
    %>

</div>
<!--#include file="includes/footer.asp"-->