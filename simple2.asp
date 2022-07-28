<%
    Dim myErr, source, stylesheet

    Set source = Server.CreateObject("Msxml2.DOMDocument.6.0") 'OLD VERSION: Microsoft.XMLDOM
    Set stylesheet = Server.CreateObject("Msxml2.DOMDocument.6.0")  'OLD VERSION: Microsoft.XMLDOM
    
    ' Load data.
    source.async = False
    source.Load Server.MapPath("data/menu.xml")

    If (source.parseError.errorCode <> 0) Then
        Set myErr = source.parseError
        MsgBox("You have error " & myErr.reason)
    Else
        ' Load style sheet.
        stylesheet.async = False
        stylesheet.Load Server.MapPath("data/menu.xsl")

        If (stylesheet.parseError.errorCode <> 0) Then
            Set myErr = stylesheet.parseError
            MsgBox("You have error " & myErr.reason)
        Else
            ' Do the transform.
            MsgBox source.transformNode(stylesheet)
        End If
    End If

    Function MsgBox(str)
        Response.Write(str)
    End Function
%>