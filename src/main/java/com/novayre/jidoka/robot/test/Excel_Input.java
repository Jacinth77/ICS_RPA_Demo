package com.novayre.jidoka.robot.test;


import java.io.Serializable;

/**
 * POJO class representing an Excel row.
 *
 * @author Jidoka
 */
public class Excel_Input implements Serializable {

    private static final long serialVersionUID = 1L;

    private String Input_ID;

    private String Status;



    public String getInput_ID() {
        return Input_ID;
    }

    public void setInput_ID(String input_id) {
        Input_ID = input_id;
    }

    public String getStatus() {
        return Status;
    }

    public void setStatus(String status) {
        Status = status;
    }

}





