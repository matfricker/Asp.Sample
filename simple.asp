<%
  'set the source and style sheet locations here.
  Dim sourceFile, styleFile
  
  sourceFile = Server.MapPath("data/menu.xml")
  styleFile = Server.MapPath("data/menu.xsl")
  
  'load the XML.
  Dim source 
  Set source = Server.CreateObject("Msxml2.DOMDocument.6.0")
  source.async = false
  source.load(sourceFile)

  'load the XSL.
  Dim style 
  Set style = Server.CreateObject("Msxml2.DOMDocument.6.0")
  style.async = false
  style.load(styleFile)

  Response.Write(source.transformNode(style))
%>