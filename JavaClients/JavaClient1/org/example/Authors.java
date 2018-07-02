/**
 * Authors.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis WSDL2Java emitter.
 */

package org.example;

public class Authors  implements java.io.Serializable {
    private java.lang.String au_Id;
    private java.lang.String au_Lname;
    private java.lang.String au_Fname;
    private java.lang.String phone;
    private java.lang.String address;
    private java.lang.String city;
    private java.lang.String state;
    private java.lang.String zip;
    private boolean contract;

    public Authors() {
    }

    public java.lang.String getAu_Id() {
        return au_Id;
    }

    public void setAu_Id(java.lang.String au_Id) {
        this.au_Id = au_Id;
    }

    public java.lang.String getAu_Lname() {
        return au_Lname;
    }

    public void setAu_Lname(java.lang.String au_Lname) {
        this.au_Lname = au_Lname;
    }

    public java.lang.String getAu_Fname() {
        return au_Fname;
    }

    public void setAu_Fname(java.lang.String au_Fname) {
        this.au_Fname = au_Fname;
    }

    public java.lang.String getPhone() {
        return phone;
    }

    public void setPhone(java.lang.String phone) {
        this.phone = phone;
    }

    public java.lang.String getAddress() {
        return address;
    }

    public void setAddress(java.lang.String address) {
        this.address = address;
    }

    public java.lang.String getCity() {
        return city;
    }

    public void setCity(java.lang.String city) {
        this.city = city;
    }

    public java.lang.String getState() {
        return state;
    }

    public void setState(java.lang.String state) {
        this.state = state;
    }

    public java.lang.String getZip() {
        return zip;
    }

    public void setZip(java.lang.String zip) {
        this.zip = zip;
    }

    public boolean isContract() {
        return contract;
    }

    public void setContract(boolean contract) {
        this.contract = contract;
    }

    private Object __equalsCalc = null;
    public synchronized boolean equals(Object obj) {
        if (!(obj instanceof Authors)) return false;
        Authors other = (Authors) obj;
        if (obj == null) return false;
        if (this == obj) return true;
        if (__equalsCalc != null) {
            return (__equalsCalc == obj);
        }
        __equalsCalc = obj;
        boolean _equals;
        _equals = true && 
            ((au_Id==null && other.getAu_Id()==null) || 
             (au_Id!=null &&
              au_Id.equals(other.getAu_Id()))) &&
            ((au_Lname==null && other.getAu_Lname()==null) || 
             (au_Lname!=null &&
              au_Lname.equals(other.getAu_Lname()))) &&
            ((au_Fname==null && other.getAu_Fname()==null) || 
             (au_Fname!=null &&
              au_Fname.equals(other.getAu_Fname()))) &&
            ((phone==null && other.getPhone()==null) || 
             (phone!=null &&
              phone.equals(other.getPhone()))) &&
            ((address==null && other.getAddress()==null) || 
             (address!=null &&
              address.equals(other.getAddress()))) &&
            ((city==null && other.getCity()==null) || 
             (city!=null &&
              city.equals(other.getCity()))) &&
            ((state==null && other.getState()==null) || 
             (state!=null &&
              state.equals(other.getState()))) &&
            ((zip==null && other.getZip()==null) || 
             (zip!=null &&
              zip.equals(other.getZip()))) &&
            contract == other.isContract();
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
        if (getAu_Id() != null) {
            _hashCode += getAu_Id().hashCode();
        }
        if (getAu_Lname() != null) {
            _hashCode += getAu_Lname().hashCode();
        }
        if (getAu_Fname() != null) {
            _hashCode += getAu_Fname().hashCode();
        }
        if (getPhone() != null) {
            _hashCode += getPhone().hashCode();
        }
        if (getAddress() != null) {
            _hashCode += getAddress().hashCode();
        }
        if (getCity() != null) {
            _hashCode += getCity().hashCode();
        }
        if (getState() != null) {
            _hashCode += getState().hashCode();
        }
        if (getZip() != null) {
            _hashCode += getZip().hashCode();
        }
        _hashCode += new Boolean(isContract()).hashCode();
        __hashCodeCalc = false;
        return _hashCode;
    }

    // Type metadata
    private static org.apache.axis.description.TypeDesc typeDesc =
        new org.apache.axis.description.TypeDesc(Authors.class);

    static {
        org.apache.axis.description.FieldDesc field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("au_Id");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "au_id"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("au_Lname");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "au_lname"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("au_Fname");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "au_fname"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("phone");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "phone"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("address");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "address"));
        field.setMinOccursIs0(true);
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("city");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "city"));
        field.setMinOccursIs0(true);
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("state");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "state"));
        field.setMinOccursIs0(true);
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("zip");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "zip"));
        field.setMinOccursIs0(true);
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("contract");
        field.setXmlName(new javax.xml.namespace.QName("http://example.org/dataset", "contract"));
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
