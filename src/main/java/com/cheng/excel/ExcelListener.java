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
import java.util.concurrent.atomic.AtomicLong;

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

	public static int getDaysOfMonth(Date date) {
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(date);
		return calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
	}


	@Override
	public void doAfterAllAnalysed(AnalysisContext context) {
		if (context.readSheetHolder().getSheetNo() == 1) {
			BigDecimal costSum = BigDecimal.ZERO;
			Map<String, Map<String, BigDecimal>> categorySum = new HashMap<>(dateCosts.size());
			List<BigDecimal> allotMoneyList = new ArrayList<>();
			projectCosts.forEach(e -> {
				e.setStartDateInit(e.getStartDate());
				e.setEndDateInit(e.getEndDate());
				if ("y".equals(HandleExcel.mode)) {
					e.setStartDate(DateUtils.truncate(e.getStartDate(), Calendar.MONTH));
					e.setEndDate(DateUtils.truncate(e.getEndDate(), Calendar.MONTH));
				}
				allotMoneyList.add(BigDecimal.ZERO);
				categorySum.put(e.getProject(), new HashMap<>());
			});
			dateCosts.forEach(dateCost -> {
				dateCost.setCostDateInit(dateCost.getCostDate());
				if ("y".equals(HandleExcel.mode)) {
					dateCost.setCostDate(DateUtils.truncate(dateCost.getCostDate(), Calendar.MONTH));
				}
			});
			List<List<Object>> sumList = new ArrayList<>();
			dateCosts.sort(Comparator.comparing(DateCost::getCostDate));
			for (DateCost dateCost : dateCosts) {
				costSum = costSum.add(dateCost.getCost());
				List<Object> list = new ArrayList<>();
				list.add(DateUtils.formatDate(dateCost.getCostDateInit(), "yyyy-MM-dd"));
				list.add(dateCost.getCategory());
				list.add(dateCost.getCost());
				for (int i = 0; i < projectCosts.size(); i++) {
					ProjectCost projectCost = projectCosts.get(i);
					Map<String, BigDecimal> key = categorySum.get(projectCost.getProject());
					BigDecimal cast = key.get(dateCost.getCategory());
					BigDecimal allot = BigDecimal.ZERO;
					if (dateCost.getCostDate().compareTo(projectCost.getStartDate()) >= 0
							&& dateCost.getCostDate().compareTo(projectCost.getEndDate()) <= 0) {
						AtomicInteger hasNumSum = new AtomicInteger(projectCost.getPersonnelNum());
						long daysOfMonth = getDaysOfMonth(dateCost.getCostDate());
						if (dateCost.getCostDate().getTime() - projectCost.getStartDate().getTime() == 0) {
							long distanceOfTwoDate = DateUtils
									.getDistanceOfTwoDate(projectCost.getStartDateInit(), dateCost.getCostDate());
							distanceOfTwoDate = distanceOfTwoDate + daysOfMonth;
							if (distanceOfTwoDate >= 0) {
								daysOfMonth = distanceOfTwoDate;
							}
						}
						if (dateCost.getCostDate().getTime() - projectCost.getEndDate().getTime() == 0) {
							long distanceOfTwoDate = DateUtils
									.getDistanceOfTwoDate(dateCost.getCostDate(), projectCost.getEndDateInit());
							if (distanceOfTwoDate >= 0) {
								daysOfMonth = distanceOfTwoDate + 1;
							}
						}
						AtomicLong hasDaysSum = new AtomicLong(daysOfMonth);
						for (int j = 0; j < projectCosts.size(); j++) {
							ProjectCost projectCostTemp = projectCosts.get(j);
							if (!projectCost.getProject().equals(projectCostTemp.getProject())) {
								if (dateCost.getCostDate().compareTo(projectCostTemp.getStartDate()) >= 0
										&& dateCost.getCostDate().compareTo(projectCostTemp.getEndDate()) <= 0) {
									hasNumSum.addAndGet(projectCostTemp.getPersonnelNum());
									long daysOfMonthTemp = getDaysOfMonth(dateCost.getCostDate());
									if (dateCost.getCostDate().getTime() - projectCostTemp.getStartDate().getTime()
											== 0) {
										long distanceOfTwoDate = DateUtils
												.getDistanceOfTwoDate(projectCostTemp.getStartDateInit(),
														dateCost.getCostDate());
										distanceOfTwoDate = distanceOfTwoDate + daysOfMonthTemp;
										if (distanceOfTwoDate >= 0) {
											daysOfMonthTemp = distanceOfTwoDate;
										}
									}
									if (dateCost.getCostDate().getTime() - projectCostTemp.getEndDate().getTime()
											== 0) {
										long distanceOfTwoDate = DateUtils.getDistanceOfTwoDate(dateCost.getCostDate(),
												projectCostTemp.getEndDateInit());
										if (distanceOfTwoDate >= 0) {
											daysOfMonthTemp = distanceOfTwoDate + 1;
										}
									}
									hasDaysSum.addAndGet(daysOfMonthTemp);
								}
							}
						}
						BigDecimal rate = new BigDecimal(projectCost.getPersonnelNum())
								.divide(new BigDecimal(hasNumSum.get()), 8, BigDecimal.ROUND_HALF_UP)
								.multiply(new BigDecimal("0.5"));
						BigDecimal daysRate = new BigDecimal(daysOfMonth)
								.divide(new BigDecimal(hasDaysSum.get()), 8, BigDecimal.ROUND_HALF_UP)
								.multiply(new BigDecimal("0.5"));
						BigDecimal add = rate.add(daysRate);
						list.add(add);
						allot = dateCost.getCost().multiply(add);
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
