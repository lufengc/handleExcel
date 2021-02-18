/*
 * Copyright &copy; cc All rights reserved.
 */

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;

/**
 *
 * @author fengcheng
 * @version 2019-06-15
 */
@Data
@ColumnWidth(15)
public class ExcelDto {

	@ExcelProperty(value = "流水号")
	private String no;

	@ExcelProperty(value = "备注")
	private String remark;

	@ExcelProperty(value = "日期")
	private String date;


}
