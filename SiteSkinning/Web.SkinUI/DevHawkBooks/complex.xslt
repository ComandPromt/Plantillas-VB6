<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

	<xsl:key name="title_type" match="title" use="@Type"/>
	
	<xsl:template name="TypeName">
		<xsl:param name="type">hello</xsl:param>
		<xsl:choose>
			<xsl:when test="$type = 'business'">
				Business
			</xsl:when>
			<xsl:when test="$type = 'mod_cook'">
				Modern Cooking
			</xsl:when>
			<xsl:when test="$type = 'popular_comp'">
				Popular Computing
			</xsl:when>
			<xsl:when test="$type = 'psychology'">
				Psychology
			</xsl:when>
			<xsl:when test="$type = 'trad_cook'">
				Traditional Cooking
			</xsl:when>
			<xsl:when test="$type = 'UNDECIDED'">
				Undecided
			</xsl:when>
			<xsl:otherwise>
				Unknown
			</xsl:otherwise>
		</xsl:choose>
	</xsl:template>

	
	<xsl:template match="/default">
		<HTML><body style="font-family: Arial; color:#000080" bgcolor="#FFCC99">
		<IMG src="DevHawkBooksMirror.jpg" alt="DevHawk Books"/>
		<table>
		<tr><td valign="top">
		<h3>Titles</h3>
		<xsl:for-each select="titles/title[count(. | key('title_type', @Type)[1])=1]">
			<xsl:sort select="@Type"/>
			<h5><xsl:call-template name="TypeName">
				<xsl:with-param name="type" select="@Type"/>
			</xsl:call-template></h5>
			<ul>
			<xsl:for-each select="key('title_type', @Type)">
				<xsl:sort select="@Name"/>
				<li>
					<a>
						<xsl:attribute name="href">title.xml?id=<xsl:value-of select="@ID" />
						</xsl:attribute>
						<xsl:value-of select="@Name" />
					</a>
				</li>
			</xsl:for-each>
			</ul>
		</xsl:for-each>
		</td><td valign="top">
		<h3>Authors</h3>
		<ul>
			<xsl:for-each select="authors/author">
				<xsl:sort select="@FirstName"/>

				<li>
					<a>
						<xsl:attribute name="href">author.xml?id=<xsl:value-of select="@ID" />
						</xsl:attribute>
						<xsl:value-of select="@FirstName" />
						<xsl:text> </xsl:text>
						<xsl:value-of select="@LastName" />
					</a>
				</li>
			</xsl:for-each>
		</ul>
		</td><td valign="top">
		<h3>Publishers</h3>
		<ul>
			<xsl:for-each select="publishers/publisher">
				<xsl:sort select="@Name"/>
				<li>
					<a>
						<xsl:attribute name="href">
						publisher.xml?id=<xsl:value-of select="@ID" />
						</xsl:attribute>
						<xsl:value-of select="@Name" />
					</a>
				</li>
			</xsl:for-each>
		</ul>
		</td></tr>
		</table>
		<a href="default.xml?skin=clear">Use extremely simple WebSkin</a>

		</body></HTML>
	</xsl:template>
	
	<xsl:template match="/author">
		<head><body style="font-family: Arial; color:#000080" bgcolor="#FFCC99">
		<IMG src="DevHawkBooksMirror.jpg" alt="DevHawk Books"/>
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
		<head><body style="font-family: Arial; color:#000080" bgcolor="#FFCC99">
		<IMG src="DevHawkBooksMirror.jpg" alt="DevHawk Books"/>
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
		<head><body style="font-family: Arial; color:#000080" bgcolor="#FFCC99">
		<IMG src="DevHawkBooksMirror.jpg" alt="DevHawk Books"/>
		<h2>Publisher: <xsl:value-of select="info/Name"/></h2>
		<img>
			<xsl:attribute name="src">publogo.ashx?id=<xsl:value-of select="info/ID" />
			</xsl:attribute>
		</img>
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

  