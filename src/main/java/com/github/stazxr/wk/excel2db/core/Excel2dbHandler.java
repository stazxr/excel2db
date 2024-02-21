package com.github.stazxr.wk.excel2db.core;

import com.github.stazxr.wk.excel2db.model.ConfigKey;
import com.github.stazxr.wk.excel2db.model.Param;
import com.github.stazxr.wk.excel2db.util.JdbcUtils;
import com.github.stazxr.wk.excel2db.util.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.*;

public class Excel2dbHandler {
    /**
     * 一秒
     */
    public static final long ONE_SECOND_OF_MILL = 1000L;

    public static final Param param = new Param();

    public static void readParam(String[] args) {
        for (String arg : args) {
            printLog("===>" + arg);
            if (arg.startsWith("configFile=")) {
                String value = arg.split("=")[1];
                param.setConfigFile(value);
            }

            if (arg.startsWith("excelFile=")) {
                String value = arg.split("=")[1];
                param.setExcelFile(value);
            }
        }
    }

    public static void readConfigFile() {
        String configFile = param.getConfigFile();
        if (StringUtils.isBlank(configFile)) {
            throw new RuntimeException("参数【configFile】为空");
        }

        File file = new File(configFile);
        if (!file.exists()) {
            throw new RuntimeException("配置文件不存在");
        }

        if (!configFile.endsWith(".properties")) {
            throw new RuntimeException("配置文件类型不正确");
        }

        Properties props = new Properties();
        try (BufferedInputStream bis = new BufferedInputStream(Files.newInputStream(file.toPath()))) {
            props.load(bis);
        } catch (IOException e) {
            throw new RuntimeException("配置文件读取异常", e);
        }

        // 设置环境变量
        System.setProperty(ConfigKey.colLength, props.getProperty(ConfigKey.colLength));
        System.setProperty(ConfigKey.dbUrl, props.getProperty(ConfigKey.dbUrl));
        System.setProperty(ConfigKey.dbUser, props.getProperty(ConfigKey.dbUser));
        System.setProperty(ConfigKey.dbPassword, props.getProperty(ConfigKey.dbPassword));
        System.setProperty(ConfigKey.insertSql, props.getProperty(ConfigKey.insertSql));
    }

    public static List<Map<String, String>> readExcelFile() {
        List<Map<String, String>> data = new ArrayList<>();
        String excelFile = param.getExcelFile();
        try (FileInputStream fis = new FileInputStream(excelFile); Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row cells : sheet) {
                // 第一行跳过
                int realRowNum = cells.getRowNum() + 1;
                if (realRowNum == 1) {
                    continue;
                }

                Map<String, String> rowData = new HashMap<>();
                int colLength = Integer.parseInt(System.getProperty(ConfigKey.colLength));
                for (int i = 0; i < colLength; i++) {
                    Cell cell = cells.getCell(i);
                    String cellValue = getStringCellValue(cell);
                    if (i == 0 && StringUtils.isBlank(cellValue)) {
                        // 读取结束
                        printLog("数据读取结束");
                        return data;
                    }

                    rowData.put("col" + (i + 1), cellValue);
                }

                data.add(rowData);
            }
        } catch (Exception e) {
            throw new RuntimeException("数据文件读取异常", e);
        }

        return data;
    }

    public static String insertDataToDb(List<Map<String, String>> data) {
        // 批次号
        long startTime = System.currentTimeMillis();
        String uuid = String.valueOf(startTime);

        printLog("数据长度为：" + data.size());

        // 创建数据库链接
        Connection connection = null;
        PreparedStatement ps = null;
        try {
            String dbUrl = System.getProperty(ConfigKey.dbUrl);
            String dbUser = System.getProperty(ConfigKey.dbUser);
            String dbPassword = System.getProperty(ConfigKey.dbPassword);
            connection = JdbcUtils.getConnection(dbUrl, dbUser, dbPassword);
            connection.setAutoCommit(false);
            printLog("数据库连接成功");

            // 初始化SQL
            String baseSql = System.getProperty(ConfigKey.insertSql);
            ps = connection.prepareStatement(baseSql);

            // 执行SQL
            int process = 1;
            for (Map<String, String> datum : data) {
                printLog("执行进度：[" + process++ + "/" + data.size() + "]");

                int i = 1;
                ps.setString(i, uuid);
                for (String key : datum.keySet()) {
                    String value = datum.get(key);
                    value = StringUtils.isBlank(value) ? "" : value;
                    ps.setString(++i, value);
                }

                ps.execute();
                ps.clearParameters();
            }

            connection.commit();
            printLog("数据提交成功");
        } catch (Exception e) {
            if (connection != null) {
                try {
                    connection.rollback();
                    printLog("数据回滚成功");
                } catch (SQLException ex) {
                    throw new RuntimeException("数据入库异常且回滚失败", ex);
                }
            }
            throw new RuntimeException("数据入库异常", e);
        } finally {
            printLog("释放数据库连接");
            JdbcUtils.close(connection, ps);

            long endTime = System.currentTimeMillis();
            printLog("数据入库总耗时：" + printCostTime(endTime - startTime));
        }

        return uuid;
    }

    public static void printLog(String log) {
        System.out.println(log);
    }

    public static String printCostTime(long cost) {
        if (cost >= ONE_SECOND_OF_MILL) {
            long s = cost / ONE_SECOND_OF_MILL;
            long ms = cost % ONE_SECOND_OF_MILL;
            return s + "秒" + ms + "毫秒";
        } else {
            return cost + "毫秒";
        }
    }

    private static String getStringCellValue(Cell cell) {
        if (cell != null) {
            // _NONE(-1), NUMERIC(0), STRING(1), FORMULA(2), BLANK(3), BOOLEAN(4), ERROR(5);
            String cellType = cell.getCellType().name();
            switch (cellType) {
                case "_NONE":
                    return null;
                case "NUMERIC":
                    return NumberToTextConverter.toText(cell.getNumericCellValue());
                case "STRING":
                    return cell.getStringCellValue();
                case "FORMULA":
                    return cell.getCellFormula();
                case "BLANK":
                    return "";
                case "BOOLEAN":
                    boolean booleanCellValue = cell.getBooleanCellValue();
                    return String.valueOf(booleanCellValue);
                case "ERROR":
                    byte errorCellValue = cell.getErrorCellValue();
                    return String.valueOf(errorCellValue);
            }
        }

        return "";
    }
}
