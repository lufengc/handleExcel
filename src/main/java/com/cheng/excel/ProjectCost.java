/*
 * Copyright &copy; cc All rights reserved.
 */

package com.cheng.excel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

import java.util.Date;

/**
 *
 * @author fengcheng
 * @version 2019-06-15
 */
@Data
@ColumnWidth(15)
public class ProjectCost {

	@ExcelProperty(value = "序号")
	private String no;

	@ExcelProperty(value = "项目名称")
	private String project;

	@ExcelProperty(value = "开始时间")
	private Date startDate;

	@ExcelProperty(value = "结束时间")
	private Date endDate;

	@ExcelProperty(value = "人数")
	private Integer personnelNum;

	private Date startDateInit;

	private Date endDateInit;

}
