/**
 * MortgageServiceLocator.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis WSDL2Java emitter.
 */

package mortgage_tools;

public class MortgageServiceLocator extends org.apache.axis.client.Service implements mortgage_tools.MortgageService {

    // Use to get a proxy class for MortgageServiceSoap
    private final java.lang.String MortgageServiceSoap_address = "http://localhost/mortgagecalc/service1.asmx";

    public String getMortgageServiceSoapAddress() {
        return MortgageServiceSoap_address;
    }

    public mortgage_tools.MortgageServiceSoap getMortgageServiceSoap() throws javax.xml.rpc.ServiceException {
       java.net.URL endpoint;
        try {
            endpoint = new java.net.URL(MortgageServiceSoap_address);
        }
        catch (java.net.MalformedURLException e) {
            return null; // unlikely as URL was validated in WSDL2Java
        }
        return getMortgageServiceSoap(endpoint);
    }

    public mortgage_tools.MortgageServiceSoap getMortgageServiceSoap(java.net.URL portAddress) throws javax.xml.rpc.ServiceException {
        try {
            return new mortgage_tools.MortgageServiceSoapStub(portAddress, this);
        }
        catch (org.apache.axis.AxisFault e) {
            return null; // ???
        }
    }

    /**
     * For the given interface, get the stub implementation.
     * If this service has no port for the given interface,
     * then ServiceException is thrown.
     */
    public java.rmi.Remote getPort(Class serviceEndpointInterface) throws javax.xml.rpc.ServiceException {
        try {
            if (mortgage_tools.MortgageServiceSoap.class.isAssignableFrom(serviceEndpointInterface)) {
                return new mortgage_tools.MortgageServiceSoapStub(new java.net.URL(MortgageServiceSoap_address), this);
            }
        }
        catch (Throwable t) {
            throw new javax.xml.rpc.ServiceException(t);
        }
        throw new javax.xml.rpc.ServiceException("There is no stub implementation for the interface:  " + (serviceEndpointInterface == null ? "null" : serviceEndpointInterface.getName()));
    }

}
