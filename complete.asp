<div class="content">
    <h3>Complete</h3>
    <p>Process Complete</p>

    <%
        If Session("strAccountNumber") <> "" Then
            Dim str 
            str1 = "<p>Account Number: AC-" + Session("strAccountNumber") + "-XXXX</p>"
            Response.Write(str1)

            str2 = "<p>Transaction Type: " + Session("strTransactionType") + "</p>"
            Response.Write(str2)
        End If
    %>

</div>