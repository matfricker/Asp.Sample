<%
    Dim pageTitle

    Dim dtmHour
    Dim strGreeting

    dtmHour = Hour(Now())
    strGreeting = "Good Evening!"

    If dtmHour <= 12 Then
        strGreeting = "Good Morning!"
    ElseIf dtmHour > 12 And dtmHour < 17 Then
        strGreeting = "Good Afternoon!"
    End If

    Sub getMessage()
        Response.Write("String from include file.")
    End Sub
%>
