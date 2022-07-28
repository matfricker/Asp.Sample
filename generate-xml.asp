<!--#include file="includes/functions.asp"-->
<% pageTitle = "Generate XML" %>
<!--#include file="includes/header.asp"-->

<%

Dim message, path
path = Server.MapPath("data")

'these are our variables
Dim objXML , objBodyItem, objPI
'create an instance of the DOM
Set objXML = Server.CreateObject("Msxml2.DOMDocument.6.0") 'Microsoft.XMLDOM
Set objPI = objXML.createProcessingInstruction("xml", "version='1.0'")
objXML.insertBefore objPI, objXML.childNodes(0)

'Create our root element using the createElement method
Set objXML.documentElement = objXML.createElement("root")

'Create the bodyitem element
Set objBodyItem = objXML.createElement("article")

'now we will create all the child elements
objBodyItem.appendChild objXML.createElement("name")
objBodyItem.appendChild objXML.createElement("description")
objBodyItem.appendChild objXML.createElement("created")

'now we add values to the child elements
objBodyItem.childNodes(0).text = "Document 1"
objBodyItem.childNodes(1).text = "Document 1 Description"
objBodyItem.childNodes(2).text = Date & " " & Time
'add the bodyitem element to the news element
objXML.documentElement.appendChild objBodyItem.cloneNode(true)
'write the document using the xml method of the DOM

'Response.Write objXML.xml
objXML.save path & "\generated-xml.xml"
message = "XML Generated"
%>

<div class="content">
    <p>
        <a target="_blank" href="data\generated-xml.xml"><% Response.Write(message) %></a>
    </p>
</div>

<!--#include file="includes/footer.asp"-->