/**
 * AuthorSetType.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis WSDL2Java emitter.
 */

package org.example;

public class AuthorSetType  implements java.io.Serializable {
    private org.example.AuthorType[] authors;

    public AuthorSetType() {
    }

    public org.example.AuthorType[] getAuthors() {
        return authors;
    }

    public void setAuthors(org.example.AuthorType[] authors) {
        this.authors = authors;
    }

    public org.example.AuthorType getAuthors(int i) {
        return authors[i];
    }

    public void setAuthors(int i, org.example.AuthorType value) {
        this.authors[i] = value;
    }

    private Object __equalsCalc = null;
    public synchronized boolean equals(Object obj) {
        if (!(obj instanceof AuthorSetType)) return false;
        AuthorSetType other = (AuthorSetType) obj;
        if (obj == null) return false;
        if (this == obj) return true;
        if (__equalsCalc != null) {
            return (__equalsCalc == obj);
        }
        __equalsCalc = obj;
        boolean _equals;
        _equals = true && 
            ((authors==null && other.getAuthors()==null) || 
             (authors!=null &&
              java.util.Arrays.equals(authors, other.getAuthors())));
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
        if (getAuthors() != null) {
            for (int i=0;
                 i<java.lang.reflect.Array.getLength(getAuthors());
                 i++) {
                Object obj = java.lang.reflect.Array.get(getAuthors(), i);
                if (obj != null &&
                    !obj.getClass().isArray()) {
                    _hashCode += obj.hashCode();
                }
            }
        }
        __hashCodeCalc = false;
        return _hashCode;
    }

    // Type metadata
    private static org.apache.axis.description.TypeDesc typeDesc =
        new org.apache.axis.description.TypeDesc(AuthorSetType.class);

    static {
        org.apache.axis.description.FieldDesc field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("authors");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "authors"));
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
