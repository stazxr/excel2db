package com.github.stazxr.wk.excel2db.util;

import com.github.stazxr.wk.excel2db.core.Excel2dbHandler;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

public class JdbcUtils {
    static {
        try {
            Class.forName("com.mysql.jdbc.Driver");
            // Class.forName("com.mysql.cj.jdbc.Driver");
        } catch (ClassNotFoundException e) {
            throw new RuntimeException("数据库驱动不存在");
        }
    }

    public static Connection getConnection(String dbUrl, String dbUser, String dbPassword) throws SQLException {
        return DriverManager.getConnection(dbUrl, dbUser, dbPassword);
    }

    public static void close(Connection connection, Statement statement) {
        try {
            if (statement != null) statement.close();
        }catch (SQLException e) {
            Excel2dbHandler.printLog("数据连接【statement】关闭失败");
        }

        try {
            if (connection != null) connection.close();
        }catch (SQLException e) {
            Excel2dbHandler.printLog("数据连接【connection】关闭失败");
        }
    }
}
