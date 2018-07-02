/*
 * Created by IntelliJ IDEA.
 * User: Administrator
 * Date: Nov 12, 2002
 * Time: 2:36:15 AM
 * To change template for new class use
 * Code Style | Class Templates options (Tools | IDE Options).
 */
import java.io.*;
import org.example.*;
import org.w3c.dom.*;
import org.apache.xerces.dom.*;
import org.apache.axis.message.*;

public class AuthorsClient {
    static void main(String[] args) throws Exception
    {
        DataSetServiceSoapStub stub = new DataSetServiceSoapStub(new java.net.URL("http://localhost/datasetservice/datasetservice.asmx"), null);
        GetAuthorsAsTypedDataSetResult result = stub.getAuthorsAsTypedDataSet();
        Object any = result.getAny();
        Element docElement = (Element)any;
        NodeList authors = docElement.getElementsByTagNameNS("http://example.org/dataset", "authors");
        for (int i=0; i<authors.getLength(); i++)
        {
            Element authorsElem = (Element)authors.item(i);
            ElementImpl fnameElem = (ElementImpl)authorsElem.getElementsByTagNameNS("http://example.org/dataset", "au_fname").item(0);
            ElementImpl lnameElem = (ElementImpl)authorsElem.getElementsByTagNameNS("http://example.org/dataset", "au_lname").item(0);
            System.out.println(fnameElem.getTextContent() + " " + lnameElem.getTextContent());
        }
    }
}
