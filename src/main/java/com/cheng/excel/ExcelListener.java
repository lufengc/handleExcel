/*
 * Copyright &copy; cc All rights reserved.
 */

package com.cheng.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.cheng.common.util.DateUtils;
import lombok.extern.slf4j.Slf4j;

import java.math.BigDecimal;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 *
 * @author fengcheng
 * @version 2019-06-15
 */
@Slf4j
public class ExcelListener extends AnalysisEventListener {

	private static List<ProjectCost> projectCosts = new ArrayList<>();

	private static List<DateCost> dateCosts = new ArrayList<>();

	@Override
	public void invoke(Object object, AnalysisContext context) {
		if (context.readSheetHolder().getSheetNo() == 0) {
			projectCosts.add((ProjectCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 1) {
			dateCosts.add((DateCost) object);
		}
	}

	@Override
	public void doAfterAllAnalysed(AnalysisContext context) {
		if (context.readSheetHolder().getSheetNo() == 1) {
			BigDecimal costSum = BigDecimal.ZERO;
			Map<String, Map<String, BigDecimal>> categorySum = new HashMap<>(dateCosts.size());
			List<BigDecimal> allotMoneyList = new ArrayList<>();
			projectCosts.forEach(e -> {
				allotMoneyList.add(BigDecimal.ZERO);
				categorySum.put(e.getProject(), new HashMap<>());
			});
			List<List<Object>> sumList = new ArrayList<>();
			dateCosts.sort(Comparator.comparing(DateCost::getCostDate));
			for (DateCost dateCost : dateCosts) {
				costSum = costSum.add(dateCost.getCost());
				List<Object> list = new ArrayList<>();
				list.add(DateUtils.formatDate(dateCost.getCostDate(), "yyyy-MM-dd"));
				list.add(dateCost.getCategory());
				list.add(dateCost.getCost());
				if ("y".equals(HandleExcel.mode)) {
					dateCost.setCostDate(DateUtils.truncate(dateCost.getCostDate(), Calendar.MONTH));
				}
				for (int i = 0; i < projectCosts.size(); i++) {
					ProjectCost projectCost = projectCosts.get(i);
					if ("y".equals(HandleExcel.mode)) {
						projectCost.setStartDate(DateUtils.truncate(projectCost.getStartDate(), Calendar.MONTH));
						projectCost.setEndDate(DateUtils.truncate(projectCost.getEndDate(), Calendar.MONTH));
					}
					Map<String, BigDecimal> key = categorySum.get(projectCost.getProject());
					BigDecimal cast = key.get(dateCost.getCategory());
					BigDecimal allot = BigDecimal.ZERO;
					if (dateCost.getCostDate().compareTo(projectCost.getStartDate()) >= 0
							&& dateCost.getCostDate().compareTo(projectCost.getEndDate()) <= 0) {
						AtomicInteger hasNumSum = new AtomicInteger(projectCost.getPersonnelNum());
						projectCosts.forEach(g -> {
							if ("y".equals(HandleExcel.mode)) {
								g.setStartDate(DateUtils.truncate(g.getStartDate(), Calendar.MONTH));
								g.setEndDate(DateUtils.truncate(g.getEndDate(), Calendar.MONTH));
							}
							if (!projectCost.getProject().equals(g.getProject())) {
								if (dateCost.getCostDate().compareTo(g.getStartDate()) >= 0
										&& dateCost.getCostDate().compareTo(g.getEndDate()) <= 0) {
									hasNumSum.addAndGet(g.getPersonnelNum());
								}
							}
						});
						BigDecimal rate = new BigDecimal(projectCost.getPersonnelNum())
								.divide(new BigDecimal(hasNumSum.get()), 8, BigDecimal.ROUND_HALF_UP);
						list.add(rate);
						allot = dateCost.getCost().multiply(rate);
						list.add(allot);
						allotMoneyList.set(i, allotMoneyList.get(i).add(allot));
					} else {
						list.add(BigDecimal.ZERO);
						list.add(BigDecimal.ZERO);
					}
					if (cast == null) {
						key.put(dateCost.getCategory(), allot);
					} else {
						key.put(dateCost.getCategory(), cast.add(allot));
					}

				}
				sumList.add(list);
			}
			List<Object> list = new ArrayList<>();
			list.add("合计");
			list.add(null);
			list.add(costSum);
			allotMoneyList.forEach(e -> {
				list.add(null);
				list.add(e);
			});
			sumList.add(list);
			writerExcle(sumList);

			writerExcleResult(categorySum);
		}
	}

	private void writerExcle(List<List<Object>> data) {
		// 动态添加 表头 headList --> 所有表头行集合
		List<List<String>> headList = new ArrayList<>();
		List<String> headTitle = new ArrayList<>();
		headTitle.add("日期");
		headTitle.add("日期");
		headList.add(headTitle);

		headTitle = new ArrayList<>();
		headTitle.add("分类");
		headTitle.add("分类");
		headList.add(headTitle);

		headTitle = new ArrayList<>();
		headTitle.add("金额");
		headTitle.add("金额");
		headList.add(headTitle);

		for (ProjectCost projectCost : projectCosts) {
			headTitle = new ArrayList<>();
			headTitle.add(projectCost.getProject());
			headTitle.add("分配率");
			headList.add(headTitle);
			headTitle = new ArrayList<>();
			headTitle.add(projectCost.getProject());
			headTitle.add("分配金额");
			headList.add(headTitle);
		}
		EasyExcel.write(HandleExcel.outPath1)
				// sheet
				.sheet("分配过程")
				// head
				.head(headList).registerWriteHandler(new ColumnWidthHandle()).doWrite(data);
	}

	private void writerExcleResult(Map<String, Map<String, BigDecimal>> categorySum) {
		// 动态添加 表头 headList --> 所有表头行集合
		List<List<String>> headList = new ArrayList<>();
		List<String> headTitle = new ArrayList<>();
		headTitle.add("序号");
		headTitle.add("序号");
		headList.add(headTitle);

		headTitle = new ArrayList<>();
		headTitle.add("项目名称");
		headTitle.add("项目名称");
		headList.add(headTitle);

		headTitle = new ArrayList<>();
		headTitle.add("开始时间");
		headTitle.add("开始时间");
		headList.add(headTitle);

		headTitle = new ArrayList<>();
		headTitle.add("结束时间");
		headTitle.add("结束时间");
		headList.add(headTitle);

		headTitle = new ArrayList<>();
		headTitle.add("人数");
		headTitle.add("人数");
		headList.add(headTitle);

		int categoryCount = 1;
		String anyKey = categorySum.keySet().stream().findAny().orElse(null);
		for (String key : categorySum.get(anyKey).keySet()) {
			headTitle = new ArrayList<>();
			headTitle.add("分配结果");
			headTitle.add(key);
			headList.add(headTitle);
			categoryCount++;
		}

		headTitle = new ArrayList<>();
		headTitle.add("合计");
		headTitle.add("合计");
		headList.add(headTitle);

		List<List<Object>> datas = new ArrayList<>();
		for (ProjectCost projectCost : projectCosts) {
			List<Object> data = new ArrayList<>();
			data.add(projectCost.getNo());
			data.add(projectCost.getProject());
			data.add(DateUtils.formatDate(projectCost.getStartDate(), "yyyy-MM-dd"));
			data.add(DateUtils.formatDate(projectCost.getEndDate(), "yyyy-MM-dd"));
			data.add(projectCost.getPersonnelNum());
			BigDecimal sum = BigDecimal.ZERO;
			for (Map.Entry<String, BigDecimal> decimalEntry : categorySum.get(projectCost.getProject()).entrySet()) {
				data.add(decimalEntry.getValue());
				sum = sum.add(decimalEntry.getValue());
			}
			data.add(sum);
			datas.add(data);
		}
		// 最后一行的合计
		List<Object> data = new ArrayList<>();
		data.add("合计");
		data.add(null);
		data.add(null);
		data.add(null);
		data.add(null);

		List<List<Object>> currentDatas = new ArrayList<>(datas);
		int startColumn = 5;
		for (int i = 0; i < categoryCount; i++) {
			BigDecimal columnSum = BigDecimal.ZERO;
			for (List<Object> objects : currentDatas) {
				columnSum = columnSum.add(new BigDecimal(objects.get(startColumn).toString()));
			}
			data.add(columnSum);
			startColumn++;
		}
		datas.add(data);

		EasyExcel.write(HandleExcel.outPath)
				// sheet
				.sheet("分配结果")
				// head
				.head(headList).registerWriteHandler(new ColumnWidthHandle()).doWrite(datas);
	}
}
