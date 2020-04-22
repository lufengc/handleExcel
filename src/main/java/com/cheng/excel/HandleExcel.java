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

	public static void main(String[] args) {
		if (args.length < 2) {
			log.error("缺少文件路径参数");
			return;
		}

		log.info("args[] = {}, {}", args[0], args[1]);

		inPath = args[0];
		outPath = args[1];

		ExcelReader excelReader = EasyExcel.read(inPath).build();
		ReadSheet readSheet1 = EasyExcel.readSheet(0).head(ProjectCost.class).registerReadListener(new ExcelListener())
				.build();
		ReadSheet readSheet2 = readSheet(1);
		ReadSheet readSheet3 = readSheet(2);
		ReadSheet readSheet4 = readSheet(3);
		ReadSheet readSheet5 = readSheet(4);
		ReadSheet readSheet6 = readSheet(5);
		ReadSheet readSheet7 = readSheet(6);
		ReadSheet readSheet8 = readSheet(7);
		ReadSheet readSheet9 = readSheet(8);
		ReadSheet readSheet10 = readSheet(9);
		excelReader.read(readSheet1, readSheet2, readSheet3, readSheet4, readSheet5, readSheet6, readSheet7, readSheet8,
				readSheet9, readSheet10);
		excelReader.finish();

	}

	private static ReadSheet readSheet(int i) {
		return EasyExcel.readSheet(i).head(DateCost.class).registerReadListener(new ExcelListener()).build();
	}
}
