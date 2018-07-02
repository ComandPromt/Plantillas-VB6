/**
 * MortgagePayments.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis WSDL2Java emitter.
 */

package mortgage_tools;

public class MortgagePayments  implements java.io.Serializable {
    private double monthlyPI;
    private double monthlyTax;
    private double monthlyInsurance;
    private double monthlyTotal;

    public MortgagePayments() {
    }

    public double getMonthlyPI() {
        return monthlyPI;
    }

    public void setMonthlyPI(double monthlyPI) {
        this.monthlyPI = monthlyPI;
    }

    public double getMonthlyTax() {
        return monthlyTax;
    }

    public void setMonthlyTax(double monthlyTax) {
        this.monthlyTax = monthlyTax;
    }

    public double getMonthlyInsurance() {
        return monthlyInsurance;
    }

    public void setMonthlyInsurance(double monthlyInsurance) {
        this.monthlyInsurance = monthlyInsurance;
    }

    public double getMonthlyTotal() {
        return monthlyTotal;
    }

    public void setMonthlyTotal(double monthlyTotal) {
        this.monthlyTotal = monthlyTotal;
    }

    private Object __equalsCalc = null;
    public synchronized boolean equals(Object obj) {
        if (!(obj instanceof MortgagePayments)) return false;
        MortgagePayments other = (MortgagePayments) obj;
        if (obj == null) return false;
        if (this == obj) return true;
        if (__equalsCalc != null) {
            return (__equalsCalc == obj);
        }
        __equalsCalc = obj;
        boolean _equals;
        _equals = true && 
            monthlyPI == other.getMonthlyPI() &&
            monthlyTax == other.getMonthlyTax() &&
            monthlyInsurance == other.getMonthlyInsurance() &&
            monthlyTotal == other.getMonthlyTotal();
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
        _hashCode += new Double(getMonthlyPI()).hashCode();
        _hashCode += new Double(getMonthlyTax()).hashCode();
        _hashCode += new Double(getMonthlyInsurance()).hashCode();
        _hashCode += new Double(getMonthlyTotal()).hashCode();
        __hashCodeCalc = false;
        return _hashCode;
    }

    // Type metadata
    private static org.apache.axis.description.TypeDesc typeDesc =
        new org.apache.axis.description.TypeDesc(MortgagePayments.class);

    static {
        org.apache.axis.description.FieldDesc field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("monthlyPI");
        field.setXmlName(new javax.xml.namespace.QName("urn:mortgage-tools", "MonthlyPI"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("monthlyTax");
        field.setXmlName(new javax.xml.namespace.QName("urn:mortgage-tools", "MonthlyTax"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("monthlyInsurance");
        field.setXmlName(new javax.xml.namespace.QName("urn:mortgage-tools", "MonthlyInsurance"));
        typeDesc.addFieldDesc(field);
        field = new org.apache.axis.description.ElementDesc();
        field.setFieldName("monthlyTotal");
        field.setXmlName(new javax.xml.namespace.QName("urn:mortgage-tools", "MonthlyTotal"));
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
