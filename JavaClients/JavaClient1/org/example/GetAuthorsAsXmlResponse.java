/**
 * GetAuthorsAsXmlResponse.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis WSDL2Java emitter.
 */

package org.example;

public class GetAuthorsAsXmlResponse  implements java.io.Serializable {
    private org.example.GetAuthorsAsXmlResult getAuthorsAsXmlResult;

    public GetAuthorsAsXmlResponse() {
    }

    public org.example.GetAuthorsAsXmlResult getGetAuthorsAsXmlResult() {
        return getAuthorsAsXmlResult;
    }

    public void setGetAuthorsAsXmlResult(org.example.GetAuthorsAsXmlResult getAuthorsAsXmlResult) {
        this.getAuthorsAsXmlResult = getAuthorsAsXmlResult;
    }

    private Object __equalsCalc = null;
    public synchronized boolean equals(Object obj) {
        if (!(obj instanceof GetAuthorsAsXmlResponse)) return false;
        GetAuthorsAsXmlResponse other = (GetAuthorsAsXmlResponse) obj;
        if (obj == null) return false;
        if (this == obj) return true;
        if (__equalsCalc != null) {
            return (__equalsCalc == obj);
        }
        __equalsCalc = obj;
        boolean _equals;
        _equals = true && 
            ((getAuthorsAsXmlResult==null && other.getGetAuthorsAsXmlResult()==null) || 
             (getAuthorsAsXmlResult!=null &&
              getAuthorsAsXmlResult.equals(other.getGetAuthorsAsXmlResult())));
        __equalsCalc = null;
        return _equals;
    }

    private boolean __hashCodeCalc = false;
    public synchronized int hashCode() {
        if (__hashCodeCalc) {
            return 0;
        }
        __hashCodeCalc = true;
        int _hashCode = 1;
        if (getGetAuthorsAsXmlResult() != null) {
            _hashCode += getGetAuthorsAsXmlResult().hashCode();
        }
        __hashCodeCalc = false;
        return _hashCode;
    }

    // Type metadata
    private static org.apache.axis.description.TypeDesc typeDesc =
        new org.apache.axis.description.TypeDesc(GetAuthorsAsXmlResponse.class);

    static {
        org.apache.axis.description.FieldDesc field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("getAuthorsAsXmlResult");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset-service", "GetAuthorsAsXmlResult"));
        field.setMinOccursIs0(true);
        typeDesc.addFieldDesc(field);
    };

    /**
     * Return type metadata object
     */
    public static org.apache.axis.description.TypeDesc getTypeDesc() {
        return typeDesc;
    }

    /**
     * Get Custom Serializer
     */
    public static org.apache.axis.encoding.Serializer getSerializer(
           String mechType, 
           Class _javaType,  
           javax.xml.namespace.QName _xmlType) {
        return 
          new  org.apache.axis.encoding.ser.BeanSerializer(
            _javaType, _xmlType, typeDesc);
    }

    /**
     * Get Custom Deserializer
     */
    public static org.apache.axis.encoding.Deserializer getDeserializer(
           String mechType, 
           Class _javaType,  
           javax.xml.namespace.QName _xmlType) {
        return 
          new  org.apache.axis.encoding.ser.BeanDeserializer(
            _javaType, _xmlType, typeDesc);
    }

}
