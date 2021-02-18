package com.cheng.excel;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import org.apache.poi.ss.usermodel.Cell;

import java.util.List;

/**
 * @author fengcheng
 */
public class ColumnWidthHandle extends AbstractColumnWidthStyleStrategy {
	/**
	 * the maximum column width in Excel is 255 characters
	 */
	private static final int MAX_COLUMN_WIDTH = 255;

	@Override
	protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<CellData> cellDataList, Cell cell, Head head,
			Integer relativeRowIndex, Boolean isHead) {
		if (isHead && cell.getRowIndex() != 0) {
			int columnWidth = cell.getStringCellValue().getBytes().length;
			if (columnWidth > MAX_COLUMN_WIDTH) {
				columnWidth = MAX_COLUMN_WIDTH;
			} else {
				columnWidth = columnWidth + 6;
			}
			writeSheetHolder.getSheet().setColumnWidth(cell.getColumnIndex(), columnWidth * 256);

		}
	}
}
