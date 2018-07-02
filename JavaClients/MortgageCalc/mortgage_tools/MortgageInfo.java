/**
 * MortgageInfo.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis WSDL2Java emitter.
 */

package mortgage_tools;

public class MortgageInfo  implements java.io.Serializable {
    private double amount;
    private double years;
    private double interest;
    private double annualTax;
    private double annualInsurance;

    public MortgageInfo() {
    }

    public double getAmount() {
        return amount;
    }

    public void setAmount(double amount) {
        this.amount = amount;
    }

    public double getYears() {
        return years;
    }

    public void setYears(double years) {
        this.years = years;
    }

    public double getInterest() {
        return interest;
    }

    public void setInterest(double interest) {
        this.interest = interest;
    }

    public double getAnnualTax() {
        return annualTax;
    }

    public void setAnnualTax(double annualTax) {
        this.annualTax = annualTax;
    }

    public double getAnnualInsurance() {
        return annualInsurance;
    }

    public void setAnnualInsurance(double annualInsurance) {
        this.annualInsurance = annualInsurance;
    }

    private Object __equalsCalc = null;
    public synchronized boolean equals(Object obj) {
        if (!(obj instanceof MortgageInfo)) return false;
        MortgageInfo other = (MortgageInfo) obj;
        if (obj == null) return false;
        if (this == obj) return true;
        if (__equalsCalc != null) {
            return (__equalsCalc == obj);
        }
        __equalsCalc = obj;
        boolean _equals;
        _equals = true && 
            amount == other.getAmount() &&
            years == other.getYears() &&
            interest == other.getInterest() &&
            annualTax == other.getAnnualTax() &&
            annualInsurance == other.getAnnualInsurance();
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
        _hashCode += new Double(getAmount()).hashCode();
        _hashCode += new Double(getYears()).hashCode();
        _hashCode += new Double(getInterest()).hashCode();
        _hashCode += new Double(getAnnualTax()).hashCode();
        _hashCode += new Double(getAnnualInsurance()).hashCode();
        __hashCodeCalc = false;
        return _hashCode;
    }

    // Type metadata
    private static org.apache.axis.description.TypeDesc typeDesc =
        new org.apache.axis.description.TypeDesc(MortgageInfo.class);

    static {
        org.apache.axis.description.FieldDesc field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("amount");
        field.setXmlName(new javax.xml.namespace.QName("urn:mortgage-tools", "amount"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("years");
        field.setXmlName(new javax.xml.namespace.QName("urn:mortgage-tools", "years"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("interest");
        field.setXmlName(new javax.xml.namespace.QName("urn:mortgage-tools", "interest"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("annualTax");
        field.setXmlName(new javax.xml.namespace.QName("urn:mortgage-tools", "annualTax"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("annualInsurance");
        field.setXmlName(new javax.xml.namespace.QName("urn:mortgage-tools", "annualInsurance"));
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
