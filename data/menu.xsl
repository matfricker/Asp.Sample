<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
    <html>
      <body style="font-family:Arial, sans-serif; font-size:12pt;
            background-color:#EEEEEE">
        <xsl:for-each select="breakfast-menu/food">
          <div style="background-color:teal; color:white; padding:4px">
            <span style="font-weight:bold; color:white"><xsl:value-of select="name"/></span>
            - <xsl:value-of select="price"/>
          </div>
          <div style="margin-left:20px; margin-bottom:1em; font-size:10pt">
            <xsl:value-of select="description"/>
            <span style="font-style:italic">
              (<xsl:value-of select="calories"/> calories per serving)
            </span>
        </div>
        </xsl:for-each>
    </body>
    </html>
  </xsl:template>
</xsl:stylesheet>