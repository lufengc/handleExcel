/*
 * Copyright &copy; cc All rights reserved.
 */

package com.cheng.excel;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

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

	@ExcelProperty(value = "研发期间")
	private String dateRange;

	@ExcelProperty(value = "人数")
	private Integer personnelNum;

	@ExcelProperty(value = "人员人工费用")
	private BigDecimal cost1 = new BigDecimal(0);

	@ExcelProperty(value = "直接投入费用")
	private BigDecimal cost2 = new BigDecimal(0);

	@ExcelProperty(value = "折旧费用与长期待摊费用")
	private BigDecimal cost3 = new BigDecimal(0);

	@ExcelProperty(value = "无形资产摊销")
	private BigDecimal cost4 = new BigDecimal(0);

	@ExcelProperty(value = "设计费用")
	private BigDecimal cost5 = new BigDecimal(0);

	@ExcelProperty(value = "装备调试费用与实验费用")
	private BigDecimal cost6 = new BigDecimal(0);

	@ExcelProperty(value = "其他费用")
	private BigDecimal cost7 = new BigDecimal(0);

	@ExcelProperty(value = "境内的委托外部研究开发费用")
	private BigDecimal cost8 = new BigDecimal(0);

	@ExcelProperty(value = "境外的研究开发费用")
	private BigDecimal cost9 = new BigDecimal(0);

	@ExcelIgnore
	private Date dateStart;

	@ExcelIgnore
	private Date dateEnd;
}
