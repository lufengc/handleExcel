/*
 * Copyright &copy; cc All rights reserved.
 */

package com.cheng.excel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.Date;

/**
 *
 * @author fengcheng
 * @version 2019-06-15
 */
@Data
public class DateCost {

	/**
	 * 日期
	 */
	@ExcelProperty
	private Date month;

	/**
	 * 金额
	 */
	@ExcelProperty
	private String amount;
}
