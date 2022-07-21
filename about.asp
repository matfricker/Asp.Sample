<!--#include file="includes/functions.asp"-->
<% pageTitle = "About Me" %>
<!--#include file="includes/header.asp"-->

<%
    Dim secretMessage
    secretMessage = Request.QueryString("msg")
%>

<div class="content">
    <p>
        This page is all about me.
        <%
            If secretMessage <> "" Then
                Response.Write("<strong>Secret Message:</strong>: " + secretMessage)
            Else
                Response.Write("<a href='about.asp?msg=Hello%20%World'>Click for secret message.</a>")
            End If
        %>
    </p>
</div>

<!--#include file="includes/footer.asp"-->