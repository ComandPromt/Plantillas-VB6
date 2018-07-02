/**
 * DataSetServiceLocator.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis WSDL2Java emitter.
 */

package org.example;

public class DataSetServiceLocator extends org.apache.axis.client.Service implements org.example.DataSetService {

    // Use to get a proxy class for DataSetServiceSoap
    private final java.lang.String DataSetServiceSoap_address = "http://localhost/DataSetService/DataSetService.asmx";

    public String getDataSetServiceSoapAddress() {
        return DataSetServiceSoap_address;
    }

    public org.example.DataSetServiceSoap getDataSetServiceSoap() throws javax.xml.rpc.ServiceException {
       java.net.URL endpoint;
        try {
            endpoint = new java.net.URL(DataSetServiceSoap_address);
        }
        catch (java.net.MalformedURLException e) {
            return null; // unlikely as URL was validated in WSDL2Java
        }
        return getDataSetServiceSoap(endpoint);
    }

    public org.example.DataSetServiceSoap getDataSetServiceSoap(java.net.URL portAddress) throws javax.xml.rpc.ServiceException {
        try {
            return new org.example.DataSetServiceSoapStub(portAddress, this);
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
            if (org.example.DataSetServiceSoap.class.isAssignableFrom(serviceEndpointInterface)) {
                return new org.example.DataSetServiceSoapStub(new java.net.URL(DataSetServiceSoap_address), this);
            }
        }
        catch (Throwable t) {
            throw new javax.xml.rpc.ServiceException(t);
        }
        throw new javax.xml.rpc.ServiceException("There is no stub implementation for the interface:  " + (serviceEndpointInterface == null ? "null" : serviceEndpointInterface.getName()));
    }

}
