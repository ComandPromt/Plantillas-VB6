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

public class AuthorsClient {
    static void main(String[] args) throws Exception
    {
        DataSetServiceSoapStub stub = new DataSetServiceSoapStub(new java.net.URL("http://localhost/datasetservice/datasetservice.asmx"), null);
        GetAuthorsAsXmlResult result2 = stub.getAuthorsAsXml();
        AuthorSetType aset = result2.getAuthorSet();
        AuthorType[] authors = aset.getAuthors();
        for (int i=0; i<authors.length; i++)
            System.out.println(authors[i].getAu_Fname() + " " + authors[i].getAu_Lname());
    }
}
