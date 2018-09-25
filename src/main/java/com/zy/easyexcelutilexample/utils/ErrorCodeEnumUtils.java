package com.zy.easyexcelutilexample.utils;

/**
 * @author Han.Sun
 * @since 2017年07月28日
 */
public enum ErrorCodeEnumUtils {

    ;


    private int code;
    private String message;

    ErrorCodeEnumUtils(int code, String message) {
        this.code = code;
        this.message = message;
    }


    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }
}
