/*
 * Copyright &copy; cc All rights reserved.
 */

package com.cheng.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;
import lombok.extern.slf4j.Slf4j;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 *
 * @author fengcheng
 * @version 2019-06-15
 */
@Slf4j
@SpringBootApplication
public class HandleExcel {

	public static String inPath = "exec/src.xlsx";
	public static String outPath = "exec/result.xlsx";
	public static String outPath1 = "exec/process.xlsx";
	public static String mode = "y";

	public static void main(String[] args) {
		if (args.length < 3) {
			log.error("缺少文件路径参数");
			return;
		}
		inPath = args[0];
		outPath = args[1];
		outPath1 = args[2];
		if (args.length == 4) {
			mode = args[3];
		}

		ExcelReader excelReader = EasyExcel.read(inPath).build();
		ReadSheet readSheet1 = EasyExcel.readSheet(0).head(ProjectCost.class).registerReadListener(new ExcelListener())
				.build();
		ReadSheet readSheet2 = readSheet(1);
		excelReader.read(readSheet1, readSheet2);
		excelReader.finish();

	}

	private static ReadSheet readSheet(int i) {
		return EasyExcel.readSheet(i).head(DateCost.class).registerReadListener(new ExcelListener()).build();
	}
}
