/*
 * Copyright &copy; cc All rights reserved.
 */

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.cheng.excel.HandleExcel;
import com.cheng.excel.ProjectCost;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author fengcheng
 * @version 2019-06-15
 */
@Slf4j
public class TestExcelListener extends AnalysisEventListener {

	private static List<ExcelDto> excelDtos = new ArrayList<>();

	@Override
	public void invoke(Object object, AnalysisContext context) {
		if (context.readSheetHolder().getSheetNo() == 0) {
			excelDtos.add((ExcelDto) object);
		}
	}

	@Override
	public void doAfterAllAnalysed(AnalysisContext context) {
		if (context.readSheetHolder().getSheetNo() == 0) {
			excelDtos.forEach(e -> {
				System.out.println(e);

			});

		}
	}

	private void writerExcle(List<ProjectCost> projectCosts) {
		EasyExcel.write(HandleExcel.outPath, ProjectCost.class).sheet("项目成本").doWrite(projectCosts);
	}

}
