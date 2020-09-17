package com.novayre.jidoka.robot.test;


import java.io.Serializable;

/**
 * POJO class representing an Excel row.
 *
 * @author Jidoka
 */
public class ExcelDSRow implements Serializable {

    private static final long serialVersionUID = 1L;

    private String Field_Name;

    private String Xpath;

    private String Value;

    private String Actions;

    public String getField_Name() {
        return Field_Name;
    }

    public void setField_Name(String field_Name) {
        Field_Name = field_Name;
    }

    public String getXpath() {
        return Xpath;
    }

    public void setXpath(String xpath) {
        Xpath = xpath;
    }

    public String getValue() {
        return Value;
    }

    public void setValue(String value) {
        Value = value;
    }

    public String getActions() {
        return Actions;
    }

    public void setActions(String actions) {
        Actions = actions;
    }

}





