package com.github.stazxr.wk.excel2db.model;

public final class Param {
    /**
     * 配置文件路径
     */
    private String configFile;

    /**
     * 数据文件路径
     */
    private String excelFile;

    public String getConfigFile() {
        return configFile;
    }

    public String getExcelFile() {
        return excelFile;
    }

    public void setConfigFile(String configFile) {
        this.configFile = configFile;
    }

    public void setExcelFile(String excelFile) {
        this.excelFile = excelFile;
    }
}
