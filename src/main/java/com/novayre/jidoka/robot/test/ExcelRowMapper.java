package com.novayre.jidoka.robot.test;

import org.apache.commons.lang.StringUtils;

import com.novayre.jidoka.data.provider.api.IExcel;
import com.novayre.jidoka.data.provider.api.IRowMapper;

public class ExcelRowMapper implements IRowMapper<IExcel, ExcelDSRow> {

    private static final int Field_Name_col = 0;

    /**
     * Search column header
     */
    public static final String Field_Name ="Field Name";

    /**
     * Column with the result title
     */
    private static final int Xpath_col = 1;

    /**
     * Search column header
     */
    public static final String Xpath ="Xpath";
    /**
     * Column with the result title
     */
    private static final int Value_col = 2;

    /**
     * Search column header
     */
    public static final String Value ="Value";

    private static final int Actions_col = 3;

    /**
     * Search column header
     */
    public static final String Actions ="Actions";

    @Override
    public ExcelDSRow map(IExcel data, int rowNum) {
        ExcelDSRow excel = new ExcelDSRow();
        excel.setField_Name(data.getCellValueAsString(rowNum, Field_Name_col));
        excel.setXpath(data.getCellValueAsString(rowNum, Xpath_col));
        excel.setValue(data.getCellValueAsString(rowNum, Value_col));
        excel.setActions(data.getCellValueAsString(rowNum, Actions_col));

        return isLastRow(excel) ? null : excel;
    }

    @Override
    public void update(IExcel data, int rowNum, ExcelDSRow rowData) {

    }


    @Override
    public boolean isLastRow(ExcelDSRow instance) {

        return instance == null || StringUtils.isBlank(instance.getField_Name());
    }



}

