/*
 * Copyright &copy; cc All rights reserved.
 */

package com.cheng.excel;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.math.BigDecimal;
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
	private Date costDate;

	/**
	 * 分类
	 */
	@ExcelProperty
	private String category;

	/**
	 * 金额
	 */
	@ExcelProperty
	private BigDecimal cost;

	private Date costDateInit;
}
