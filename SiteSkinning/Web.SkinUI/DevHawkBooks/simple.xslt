<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:template match="/default">
		<HTML><BODY>
		<H1>DevHawk Books</H1>
		<table>
		<tr><td valign="top">
		<h3>Authors</h3>
		<ul>
			<xsl:for-each select="authors/author">
				<li>
					<a>
						<xsl:attribute name="href">
						author.xml?id=<xsl:value-of select="@ID" />
						</xsl:attribute>
						<xsl:value-of select="@FirstName" />
						<xsl:text> </xsl:text>
						<xsl:value-of select="@LastName" />
					</a>
				</li>
			</xsl:for-each>
		</ul>
		</td><td valign="top">
		<h3>Titles</h3>
		<ul>
			<xsl:for-each select="titles/title">
				<li>
					<a>
						<xsl:attribute name="href">
						title.xml?id=<xsl:value-of select="@ID" />
						</xsl:attribute>
						<xsl:value-of select="@Name" />
					</a>
				</li>
			</xsl:for-each>
		</ul>		</td><td valign="top">
		<h3>Publishers</h3>
		<ul>
			<xsl:for-each select="publishers/publisher">
				<li>
					<a>
						<xsl:attribute name="href">
						publisher.xml?id=<xsl:value-of select="@ID" />
						</xsl:attribute>
						<xsl:value-of select="@Name" />
					</a>
				</li>
			</xsl:for-each>
		</ul>		</td></tr>
		</table>
		<a href="default.xml?skin=complex.xslt">Use slightly more exciting WebSkin</a>
		</BODY></HTML>
	</xsl:template>
	
	<xsl:template match="/author">
		<head><body>
		<h1>DevHawk Books</h1>
		<h2>Author: <xsl:value-of select="info/FirstName"/><xsl:text> </xsl:text><xsl:value-of select="info/LastName"/> </h2>
		<h3>Address</h3>
		<p>
			<xsl:value-of select="info/Address"/><br/>
			<xsl:value-of select="info/City"/>, <xsl:value-of select="info/State"/><xsl:text> </xsl:text><xsl:value-of select="info/Zip"/><br/>
			<xsl:value-of select="info/Phone"/>
		</p>
		<h3>Titles</h3>
		<ul>
			<xsl:for-each select="titles/title">
				<li>
					<a>
						<xsl:attribute name="href">
						title.xml?id=<xsl:value-of select="@ID" />
						</xsl:attribute>
						<xsl:value-of select="@Name" />
					</a>
				</li>
			</xsl:for-each>
		</ul>
		<a href="default.xml">return to homepage</a>
		</body></head>
	
	</xsl:template>


	<xsl:template match="/title">
		<head><body>
		<h1>DevHawk Books</h1>
		<h2>Title: <xsl:value-of select="info/Name"/></h2>
		<p><xsl:value-of select="info/Notes"/></p>
		<h3>Authors</h3>
		<ul>
			<xsl:for-each select="authors/author">
				<li>
					<a>
						<xsl:attribute name="href">
						author.xml?id=<xsl:value-of select="@ID" />
						</xsl:attribute>
						<xsl:value-of select="@FirstName" />
						<xsl:text> </xsl:text>
						<xsl:value-of select="@LastName" />
					</a>
				</li>
			</xsl:for-each>
		</ul>
		<h3>Publisher</h3>
		<p>
			<a>
				<xsl:attribute name="href">
				publisher.xml?id=<xsl:value-of select="publisher/@ID" />
				</xsl:attribute>
				<xsl:value-of select="publisher/@Name" />
			</a>
		</p>
		<a href="default.xml">return to homepage</a>

		</body></head>
	
	</xsl:template>
	
	<xsl:template match="/publisher">
		<head><body>
		<h1>DevHawk Books</h1>
		<h2>Publisher: <xsl:value-of select="info/Name"/></h2>
		<p><xsl:value-of select="info/City"/>, <xsl:value-of select="info/State"/><xsl:text> </xsl:text><xsl:value-of select="info/Country"/><br/></p>
		<h3>Titles</h3>
		<ul>
			<xsl:for-each select="titles/title">
				<li>
					<a>
						<xsl:attribute name="href">
						title.xml?id=<xsl:value-of select="@ID" />
						</xsl:attribute>
						<xsl:value-of select="@Name" />
					</a>
				</li>
			</xsl:for-each>
		</ul>
		<a href="default.xml">return to homepage</a>
		</body></head>
	
	</xsl:template>

</xsl:stylesheet>

  