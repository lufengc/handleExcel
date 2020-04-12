/*
 * Copyright &copy; cc All rights reserved.
 */

package com.cheng.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.cheng.common.util.DateUtils;
import lombok.extern.slf4j.Slf4j;
import org.jetbrains.annotations.NotNull;

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

	private static List<DateCost> dateCosts1 = new ArrayList<>();
	private static List<DateCost> dateCosts2 = new ArrayList<>();
	private static List<DateCost> dateCosts3 = new ArrayList<>();
	private static List<DateCost> dateCosts4 = new ArrayList<>();
	private static List<DateCost> dateCosts5 = new ArrayList<>();
	private static List<DateCost> dateCosts6 = new ArrayList<>();
	private static List<DateCost> dateCosts7 = new ArrayList<>();
	private static List<DateCost> dateCosts8 = new ArrayList<>();
	private static List<DateCost> dateCosts9 = new ArrayList<>();

	@Override
	public void invoke(Object object, AnalysisContext context) {
		if (context.readSheetHolder().getSheetNo() == 0) {
			projectCosts.add((ProjectCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 1) {
			dateCosts1.add((DateCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 2) {
			dateCosts2.add((DateCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 3) {
			dateCosts3.add((DateCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 4) {
			dateCosts4.add((DateCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 5) {
			dateCosts5.add((DateCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 6) {
			dateCosts6.add((DateCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 7) {
			dateCosts7.add((DateCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 8) {
			dateCosts8.add((DateCost) object);
		} else if (context.readSheetHolder().getSheetNo() == 9) {
			dateCosts9.add((DateCost) object);
		}
	}

	@Override
	public void doAfterAllAnalysed(AnalysisContext context) {
		if (context.readSheetHolder().getSheetNo() == 9) {
			Map<Integer, BigDecimal> maps1 = getMonthMoneyMap(dateCosts1);
			Map<Integer, BigDecimal> maps2 = getMonthMoneyMap(dateCosts2);
			Map<Integer, BigDecimal> maps3 = getMonthMoneyMap(dateCosts3);
			Map<Integer, BigDecimal> maps4 = getMonthMoneyMap(dateCosts4);
			Map<Integer, BigDecimal> maps5 = getMonthMoneyMap(dateCosts5);
			Map<Integer, BigDecimal> maps6 = getMonthMoneyMap(dateCosts6);
			Map<Integer, BigDecimal> maps7 = getMonthMoneyMap(dateCosts7);
			Map<Integer, BigDecimal> maps8 = getMonthMoneyMap(dateCosts8);
			Map<Integer, BigDecimal> maps9 = getMonthMoneyMap(dateCosts9);

			projectCosts.forEach(e -> {
				String[] split = e.getDateRange().split("到");
				int first = 0;
				int second = 0;

				boolean isDay = false;
				if (split[0].length() > 8) {
					isDay = true;
				}

				Date date1 = DateUtils.parseDate(split[0]);
				Date date2 = DateUtils.parseDate(split[1]);
				first = DateUtils.toCalendar(date1).get(Calendar.MONTH);
				second = DateUtils.toCalendar(date2).get(Calendar.MONTH);

				for (int i = 0; i < 12; i++) {
					if (first <= second) {
						e.getMonthRangeArr().add(first);
					}
					first++;
				}
			});
			projectCosts.forEach(e -> e.getMonthRangeArr().forEach(r -> {
				int month = r;
				AtomicInteger hasNumSum = new AtomicInteger(e.getPersonnelNum());
				projectCosts.forEach(g -> {
					if (!Objects.equals(e.getProject(), g.getProject()) && g.getMonthRangeArr().contains(r)) {
						hasNumSum.addAndGet(g.getPersonnelNum());
					}
				});
				BigDecimal money1 = maps1.get(month);
				if (money1 != null) {
					money1 = money1.multiply(new BigDecimal(e.getPersonnelNum()))
							.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
					e.setCost1(e.getCost1().add(money1));
				}
				BigDecimal money2 = maps2.get(month);
				if (money2 != null && money2.compareTo(BigDecimal.ZERO) > 0) {
					money2 = money2.multiply(new BigDecimal(e.getPersonnelNum()))
							.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
					e.setCost2(e.getCost2().add(money2));
				}
				BigDecimal money3 = maps3.get(month);
				if (money3 != null && money3.compareTo(BigDecimal.ZERO) > 0) {
					money3 = money3.multiply(new BigDecimal(e.getPersonnelNum()))
							.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
					e.setCost3(e.getCost3().add(money3));
				}
				BigDecimal money4 = maps4.get(month);
				if (money4 != null && money4.compareTo(BigDecimal.ZERO) > 0) {
					money4 = money4.multiply(new BigDecimal(e.getPersonnelNum()))
							.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
					e.setCost4(e.getCost4().add(money4));
				}
				BigDecimal money5 = maps5.get(month);
				if (money5 != null && money5.compareTo(BigDecimal.ZERO) > 0) {
					money5 = money5.multiply(new BigDecimal(e.getPersonnelNum()))
							.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
					e.setCost5(e.getCost5().add(money5));
				}
				BigDecimal money6 = maps6.get(month);
				if (money6 != null && money6.compareTo(BigDecimal.ZERO) > 0) {
					money6 = money6.multiply(new BigDecimal(e.getPersonnelNum()))
							.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
					e.setCost6(e.getCost6().add(money6));
				}
				BigDecimal money7 = maps7.get(month);
				if (money7 != null && money7.compareTo(BigDecimal.ZERO) > 0) {
					money7 = money7.multiply(new BigDecimal(e.getPersonnelNum()))
							.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
					e.setCost7(e.getCost7().add(money7));
				}
				BigDecimal money8 = maps8.get(month);
				if (money8 != null && money8.compareTo(BigDecimal.ZERO) > 0) {
					money8 = money8.multiply(new BigDecimal(e.getPersonnelNum()))
							.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
					e.setCost8(e.getCost8().add(money8));
				}
				BigDecimal money9 = maps9.get(month);
				if (money9 != null && money9.compareTo(BigDecimal.ZERO) > 0) {
					money9 = money9.multiply(new BigDecimal(e.getPersonnelNum()))
							.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
					e.setCost9(e.getCost9().add(money9));
				}
			}));
			projectCosts.forEach(System.out::println);
			writerExcle(projectCosts);
		}
	}

	@NotNull
	private Map<Integer, BigDecimal> getMonthMoneyMap(List<DateCost> dateCosts) {
		Map<Integer, BigDecimal> maps = new HashMap<>(50);
		dateCosts.forEach(e -> {
			Date costDate = DateUtils.parseDate(e.getCostDate());
			int month = DateUtils.toCalendar(costDate).get(Calendar.MONTH);
			BigDecimal count = maps.get(month);
			if (count == null) {
				maps.put(month, e.getCost());
			} else {
				maps.put(month, e.getCost().add(count));
			}
		});
		System.out.println(maps.values());
		return maps;
	}

	private void writerExcle(List<ProjectCost> projectCosts) {
		EasyExcel.write(HandleExcel.outPath, ProjectCost.class).sheet("项目成本").doWrite(projectCosts);
	}

}
