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

    Sub consoleLog(str)
        Response.Write("<script>console.log(`'" & str & "'`);</script>")
    End Sub

    Function addLeadingZero(value)
        addLeadingZero = value
        If value < 10 then
            addLeadingZero = "0" & value
        End If
    End Function
%>
