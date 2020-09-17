package com.novayre.jidoka.robot.test;

import org.apache.commons.lang.StringUtils;

import com.novayre.jidoka.data.provider.api.IExcel;
import com.novayre.jidoka.data.provider.api.IRowMapper;

public class Excel_Input_RowMapper implements IRowMapper<IExcel, Excel_Input> {

    private static final int Input_ID_col = 0;

    /**
     * Search column header
     */
    public static final String Input_ID ="Input_ID";

    /**
     * Column with the result title
     */
    private static final int Status_col = 1;

    /**
     * Search column header
     */
    public static final String Status ="Status";
    /**
     * Column with the result title
     */
    private static final int Value_col = 2;


    @Override
    public Excel_Input map(IExcel data, int rowNum) {
        Excel_Input excel = new Excel_Input();
        excel.setField_Name(data.getCellValueAsString(rowNum, Input_ID_col));
        excel.setXpath(data.getCellValueAsString(rowNum, Status_col));


        return isLastRow(excel) ? null : excel;
    }

    @Override
    public void update(IExcel data, int rowNum, Excel_Input rowData) {

    }


    @Override
    public boolean isLastRow(Excel_Input instance) {

        return instance == null || StringUtils.isBlank(instance.getInput_ID());
    }



}

