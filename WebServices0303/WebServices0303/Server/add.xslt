<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
			   xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xsl:version="1.0">
	<soap:Body>
		<ns:AddResponse xmlns:ns="urn:msdn-microsoft-com:hows">
			<ns:sum><xsl:value-of select="sum(//ns:Add/*/text())"/></ns:sum>
		</ns:AddResponse>
	</soap:Body>
</soap:Envelope>
