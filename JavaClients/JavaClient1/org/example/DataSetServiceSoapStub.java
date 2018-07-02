/**
 * DataSetServiceSoapStub.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis WSDL2Java emitter.
 */

package org.example;

public class DataSetServiceSoapStub extends org.apache.axis.client.Stub implements org.example.DataSetServiceSoap {
    private java.util.Vector cachedSerClasses = new java.util.Vector();
    private java.util.Vector cachedSerQNames = new java.util.Vector();
    private java.util.Vector cachedSerFactories = new java.util.Vector();
    private java.util.Vector cachedDeserFactories = new java.util.Vector();

    public DataSetServiceSoapStub() throws org.apache.axis.AxisFault {
         this(null);
    }

    public DataSetServiceSoapStub(java.net.URL endpointURL, javax.xml.rpc.Service service) throws org.apache.axis.AxisFault {
         this(service);
         super.cachedEndpoint = endpointURL;
    }

    public DataSetServiceSoapStub(javax.xml.rpc.Service service) throws org.apache.axis.AxisFault {
        try {
            if (service == null) {
                super.service = new org.apache.axis.client.Service();
            } else {
                super.service = service;
            }
            Class cls;
            javax.xml.namespace.QName qName;
            Class beansf = org.apache.axis.encoding.ser.BeanSerializerFactory.class;
            Class beandf = org.apache.axis.encoding.ser.BeanDeserializerFactory.class;
            Class enumsf = org.apache.axis.encoding.ser.EnumSerializerFactory.class;
            Class enumdf = org.apache.axis.encoding.ser.EnumDeserializerFactory.class;
            Class arraysf = org.apache.axis.encoding.ser.ArraySerializerFactory.class;
            Class arraydf = org.apache.axis.encoding.ser.ArrayDeserializerFactory.class;
            Class simplesf = org.apache.axis.encoding.ser.SimpleNonPrimitiveSerializerFactory.class;
            Class simpledf = org.apache.axis.encoding.ser.SimpleDeserializerFactory.class;
            qName = new javax.xml.namespace.QName("http://example.org/dataset", ">AuthorSet>authors");
            cachedSerQNames.add(qName);
            cls = org.example.Authors.class;
            cachedSerClasses.add(cls);
            cachedSerFactories.add(beansf);
            cachedDeserFactories.add(beandf);

            qName = new javax.xml.namespace.QName("http://example.org/dataset-service", ">GetAuthorsAsTypedDataSetResponse>GetAuthorsAsTypedDataSetResult");
            cachedSerQNames.add(qName);
            cls = org.example.GetAuthorsAsTypedDataSetResult.class;
            cachedSerClasses.add(cls);
            cachedSerFactories.add(beansf);
            cachedDeserFactories.add(beandf);

            qName = new javax.xml.namespace.QName("http://example.org/dataset-service", ">GetAuthorsAsXml");
            cachedSerQNames.add(qName);
            cls = org.example.GetAuthorsAsXml.class;
            cachedSerClasses.add(cls);
            cachedSerFactories.add(beansf);
            cachedDeserFactories.add(beandf);

            qName = new javax.xml.namespace.QName("http://example.org/dataset", ">AuthorSet");
            cachedSerQNames.add(qName);
            cls = org.example.AuthorSet.class;
            cachedSerClasses.add(cls);
            cachedSerFactories.add(beansf);
            cachedDeserFactories.add(beandf);

            qName = new javax.xml.namespace.QName("http://example.org/dataset-service", ">GetAuthorsAsXmlResponse>GetAuthorsAsXmlResult");
            cachedSerQNames.add(qName);
            cls = org.example.GetAuthorsAsXmlResult.class;
            cachedSerClasses.add(cls);
            cachedSerFactories.add(beansf);
            cachedDeserFactories.add(beandf);

            qName = new javax.xml.namespace.QName("http://example.org/dataset-service", ">GetAuthorsAsXmlResponse");
            cachedSerQNames.add(qName);
            cls = org.example.GetAuthorsAsXmlResponse.class;
            cachedSerClasses.add(cls);
            cachedSerFactories.add(beansf);
            cachedDeserFactories.add(beandf);

        }
        catch(java.lang.Exception t) {
            throw org.apache.axis.AxisFault.makeFault(t);
        }
    }

    private org.apache.axis.client.Call createCall() throws java.rmi.RemoteException {
        try {
            org.apache.axis.client.Call call =
                    (org.apache.axis.client.Call) super.service.createCall();
            if (super.maintainSessionSet) {
                call.setMaintainSession(super.maintainSession);
            }
            if (super.cachedUsername != null) {
                call.setUsername(super.cachedUsername);
            }
            if (super.cachedPassword != null) {
                call.setPassword(super.cachedPassword);
            }
            if (super.cachedEndpoint != null) {
                call.setTargetEndpointAddress(super.cachedEndpoint);
            }
            if (super.cachedTimeout != null) {
                call.setTimeout(super.cachedTimeout);
            }
            java.util.Enumeration keys = super.cachedProperties.keys();
            while (keys.hasMoreElements()) {
                String key = (String) keys.nextElement();
                if(call.isPropertySupported(key))
                    call.setProperty(key, super.cachedProperties.get(key));
                else
                    call.setScopedProperty(key, super.cachedProperties.get(key));
            }
            // All the type mapping information is registered
            // when the first call is made.
            // The type mapping information is actually registered in
            // the TypeMappingRegistry of the service, which
            // is the reason why registration is only needed for the first call.
            synchronized (this) {
                if (firstCall()) {
                    // must set encoding style before registering serializers
                    call.setEncodingStyle(null);
                    for (int i = 0; i < cachedSerFactories.size(); ++i) {
                        Class cls = (Class) cachedSerClasses.get(i);
                        javax.xml.namespace.QName qName =
                                (javax.xml.namespace.QName) cachedSerQNames.get(i);
                        Class sf = (Class)
                                 cachedSerFactories.get(i);
                        Class df = (Class)
                                 cachedDeserFactories.get(i);
                        call.registerTypeMapping(cls, qName, sf, df, false);
                    }
                }
            }
            return call;
        }
        catch (Throwable t) {
            throw new org.apache.axis.AxisFault("Failure trying to get the Call object", t);
        }
    }

    public org.example.GetAuthorsAsTypedDataSetResult getAuthorsAsTypedDataSet() throws java.rmi.RemoteException{
        if (super.cachedEndpoint == null) {
            throw new org.apache.axis.NoEndPointException();
        }
        org.apache.axis.client.Call call = createCall();
        call.setReturnType(new javax.xml.namespace.QName("http://example.org/dataset-service", ">GetAuthorsAsTypedDataSetResponse>GetAuthorsAsTypedDataSetResult"));
        call.setUseSOAPAction(true);
        call.setSOAPActionURI("http://example.org/dataset-service/GetAuthorsAsTypedDataSet");
        call.setEncodingStyle(null);
        call.setScopedProperty(org.apache.axis.AxisEngine.PROP_DOMULTIREFS, Boolean.FALSE);
        call.setScopedProperty(org.apache.axis.client.Call.SEND_TYPE_ATTR, Boolean.FALSE);
        call.setOperationStyle("wrapped");
        call.setOperationName(new javax.xml.namespace.QName("http://example.org/dataset-service", "GetAuthorsAsTypedDataSet"));

        Object resp = call.invoke(new Object[] {});

        if (resp instanceof java.rmi.RemoteException) {
            throw (java.rmi.RemoteException)resp;
        }
        else {
            try {
                return (org.example.GetAuthorsAsTypedDataSetResult) resp;
            } catch (java.lang.Exception e) {
                return (org.example.GetAuthorsAsTypedDataSetResult) org.apache.axis.utils.JavaUtils.convert(resp, org.example.GetAuthorsAsTypedDataSetResult.class);
            }
        }
    }

    public org.example.GetAuthorsAsXmlResult getAuthorsAsXml() throws java.rmi.RemoteException{
        if (super.cachedEndpoint == null) {
            throw new org.apache.axis.NoEndPointException();
        }
        org.apache.axis.client.Call call = createCall();
        call.setReturnType(new javax.xml.namespace.QName("http://example.org/dataset-service", ">GetAuthorsAsXmlResponse>GetAuthorsAsXmlResult"));
        call.setUseSOAPAction(true);
        call.setSOAPActionURI("http://example.org/dataset-service/GetAuthorsAsXml");
        call.setEncodingStyle(null);
        call.setScopedProperty(org.apache.axis.AxisEngine.PROP_DOMULTIREFS, Boolean.FALSE);
        call.setScopedProperty(org.apache.axis.client.Call.SEND_TYPE_ATTR, Boolean.FALSE);
        call.setOperationStyle("wrapped");
        call.setOperationName(new javax.xml.namespace.QName("http://example.org/dataset-service", "GetAuthorsAsXml"));

        Object resp = call.invoke(new Object[] {});

        if (resp instanceof java.rmi.RemoteException) {
            throw (java.rmi.RemoteException)resp;
        }
        else {
            try {
                return (org.example.GetAuthorsAsXmlResult) resp;
            } catch (java.lang.Exception e) {
                return (org.example.GetAuthorsAsXmlResult) org.apache.axis.utils.JavaUtils.convert(resp, org.example.GetAuthorsAsXmlResult.class);
            }
        }
    }

}
