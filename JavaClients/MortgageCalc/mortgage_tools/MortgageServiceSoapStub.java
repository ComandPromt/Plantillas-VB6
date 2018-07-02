/**
 * MortgageServiceSoapStub.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis WSDL2Java emitter.
 */

package mortgage_tools;

public class MortgageServiceSoapStub extends org.apache.axis.client.Stub implements mortgage_tools.MortgageServiceSoap {
    private java.util.Vector cachedSerClasses = new java.util.Vector();
    private java.util.Vector cachedSerQNames = new java.util.Vector();
    private java.util.Vector cachedSerFactories = new java.util.Vector();
    private java.util.Vector cachedDeserFactories = new java.util.Vector();

    public MortgageServiceSoapStub() throws org.apache.axis.AxisFault {
         this(null);
    }

    public MortgageServiceSoapStub(java.net.URL endpointURL, javax.xml.rpc.Service service) throws org.apache.axis.AxisFault {
         this(service);
         super.cachedEndpoint = endpointURL;
    }

    public MortgageServiceSoapStub(javax.xml.rpc.Service service) throws org.apache.axis.AxisFault {
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
            qName = new javax.xml.namespace.QName("urn:mortgage-tools", "MortgagePayments");
            cachedSerQNames.add(qName);
            cls = mortgage_tools.MortgagePayments.class;
            cachedSerClasses.add(cls);
            cachedSerFactories.add(beansf);
            cachedDeserFactories.add(beandf);

            qName = new javax.xml.namespace.QName("urn:mortgage-tools", "MortgageInfo");
            cachedSerQNames.add(qName);
            cls = mortgage_tools.MortgageInfo.class;
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

    public mortgage_tools.MortgagePayments calculateMortgage(mortgage_tools.MortgageInfo minfo) throws java.rmi.RemoteException{
        if (super.cachedEndpoint == null) {
            throw new org.apache.axis.NoEndPointException();
        }
        org.apache.axis.client.Call call = createCall();
        javax.xml.namespace.QName p0QName = new javax.xml.namespace.QName("urn:mortgage-tools", "minfo");
        call.addParameter(p0QName, new javax.xml.namespace.QName("urn:mortgage-tools", "MortgageInfo"), mortgage_tools.MortgageInfo.class, javax.xml.rpc.ParameterMode.IN);
        call.setReturnType(new javax.xml.namespace.QName("urn:mortgage-tools", "MortgagePayments"));
        call.setUseSOAPAction(true);
        call.setSOAPActionURI("urn:mortgage-tools/CalculateMortgage");
        call.setEncodingStyle(null);
        call.setScopedProperty(org.apache.axis.AxisEngine.PROP_DOMULTIREFS, Boolean.FALSE);
        call.setScopedProperty(org.apache.axis.client.Call.SEND_TYPE_ATTR, Boolean.FALSE);
        call.setOperationStyle("wrapped");
        call.setOperationName(new javax.xml.namespace.QName("urn:mortgage-tools", "CalculateMortgage"));

        Object resp = call.invoke(new Object[] {minfo});

        if (resp instanceof java.rmi.RemoteException) {
            throw (java.rmi.RemoteException)resp;
        }
        else {
            try {
                return (mortgage_tools.MortgagePayments) resp;
            } catch (java.lang.Exception e) {
                return (mortgage_tools.MortgagePayments) org.apache.axis.utils.JavaUtils.convert(resp, mortgage_tools.MortgagePayments.class);
            }
        }
    }

    public java.lang.Object getSomething() throws java.rmi.RemoteException{
        if (super.cachedEndpoint == null) {
            throw new org.apache.axis.NoEndPointException();
        }
        org.apache.axis.client.Call call = createCall();
        call.setReturnType(new javax.xml.namespace.QName("http://www.w3.org/2001/XMLSchema", "anyType"));
        call.setUseSOAPAction(true);
        call.setSOAPActionURI("urn:mortgage-tools/GetSomething");
        call.setEncodingStyle(null);
        call.setScopedProperty(org.apache.axis.AxisEngine.PROP_DOMULTIREFS, Boolean.FALSE);
        call.setScopedProperty(org.apache.axis.client.Call.SEND_TYPE_ATTR, Boolean.FALSE);
        call.setOperationStyle("wrapped");
        call.setOperationName(new javax.xml.namespace.QName("urn:mortgage-tools", "GetSomething"));

        Object resp = call.invoke(new Object[] {});

        if (resp instanceof java.rmi.RemoteException) {
            throw (java.rmi.RemoteException)resp;
        }
        else {
            try {
                return (java.lang.Object) resp;
            } catch (java.lang.Exception e) {
                return (java.lang.Object) org.apache.axis.utils.JavaUtils.convert(resp, java.lang.Object.class);
            }
        }
    }

}
