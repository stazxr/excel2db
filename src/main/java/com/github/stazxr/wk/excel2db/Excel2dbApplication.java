package com.github.stazxr.wk.excel2db;

import com.github.stazxr.wk.excel2db.core.Excel2dbHandler;

import java.util.List;
import java.util.Map;

/**
 * 自动将 Excel 数据读取到数据库的工具包
 *
 * @author SunTao
 * @since 2023-01-18
 */
public class Excel2dbApplication {
	public static void main(String[] args) {
		// 配置参数
		Excel2dbHandler.readParam(args);

		// 设置环境变量
		Excel2dbHandler.readConfigFile();

		// 读取数据
		List<Map<String, String>> data = Excel2dbHandler.readExcelFile();

		// 数据入库
		String pcNum = Excel2dbHandler.insertDataToDb(data);
		Excel2dbHandler.printLog("数据导入成功，批次号为：" + pcNum);
	}
}
