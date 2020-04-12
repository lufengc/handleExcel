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
import java.util.concurrent.atomic.AtomicBoolean;
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

			AtomicBoolean isMonth = new AtomicBoolean(false);

			projectCosts.forEach(e -> {
				String[] split = e.getDateRange().split("到");

				if (split[0].length() < 8) {
					isMonth.set(true);
				}

				Date date1 = DateUtils.parseDate(split[0]);
				Date date2 = DateUtils.parseDate(split[1]);

				if (isMonth.get()) {
					date1 = DateUtils.parseDate(DateUtils.formatDate(date1, "yyyy-MM"));
					date2 = DateUtils.parseDate(DateUtils.formatDate(date2, "yyyy-MM"));
				}

				e.setDateStart(date1);
				e.setDateEnd(date2);

			});

			Map<Date, BigDecimal> maps1 = getDateCostMap(dateCosts1, isMonth);
			Map<Date, BigDecimal> maps2 = getDateCostMap(dateCosts2, isMonth);
			Map<Date, BigDecimal> maps3 = getDateCostMap(dateCosts3, isMonth);
			Map<Date, BigDecimal> maps4 = getDateCostMap(dateCosts4, isMonth);
			Map<Date, BigDecimal> maps5 = getDateCostMap(dateCosts5, isMonth);
			Map<Date, BigDecimal> maps6 = getDateCostMap(dateCosts6, isMonth);
			Map<Date, BigDecimal> maps7 = getDateCostMap(dateCosts7, isMonth);
			Map<Date, BigDecimal> maps8 = getDateCostMap(dateCosts8, isMonth);
			Map<Date, BigDecimal> maps9 = getDateCostMap(dateCosts9, isMonth);

			projectCosts.forEach(e -> {
				Date startDate = e.getDateStart();
				Date endate = e.getDateEnd();

				for (int i = 0; i < 1000; i++) {
					if (startDate.compareTo(endate) <= 0) {
						System.out.println(DateUtils.formatDate(startDate) + " <= " + DateUtils.formatDate(endate));
						AtomicInteger hasNumSum = new AtomicInteger(e.getPersonnelNum());
						for (ProjectCost projectCost : projectCosts) {
							if (!Objects.equals(e.getProject(), projectCost.getProject())
									&& startDate.compareTo(projectCost.getDateStart()) >= 0
									&& startDate.compareTo(projectCost.getDateEnd()) <= 0) {
								hasNumSum.addAndGet(projectCost.getPersonnelNum());
							}
						}

						BigDecimal money1 = maps1.get(startDate);
						if (money1 != null) {
							money1 = money1.multiply(new BigDecimal(e.getPersonnelNum()))
									.divide(new BigDecimal(hasNumSum.get()), 4, BigDecimal.ROUND_HALF_UP);
							e.setCost1(e.getCost1().add(money1));
						}
						BigDecimal money2 = maps2.get(startDate);
						if (money2 != null && money2.compareTo(BigDecimal.ZERO) > 0) {
							money2 = money2.multiply(new BigDecimal(e.getPersonnelNum()))
									.divide(new BigDecimal(hasNumSum.get()), 4, BigDecimal.ROUND_HALF_UP);
							e.setCost2(e.getCost2().add(money2));
						}
						BigDecimal money3 = maps3.get(startDate);
						if (money3 != null && money3.compareTo(BigDecimal.ZERO) > 0) {
							money3 = money3.multiply(new BigDecimal(e.getPersonnelNum()))
									.divide(new BigDecimal(hasNumSum.get()), 4, BigDecimal.ROUND_HALF_UP);
							e.setCost3(e.getCost3().add(money3));
						}
						BigDecimal money4 = maps4.get(startDate);
						if (money4 != null && money4.compareTo(BigDecimal.ZERO) > 0) {
							money4 = money4.multiply(new BigDecimal(e.getPersonnelNum()))
									.divide(new BigDecimal(hasNumSum.get()), 4, BigDecimal.ROUND_HALF_UP);
							e.setCost4(e.getCost4().add(money4));
						}
						BigDecimal money5 = maps5.get(startDate);
						if (money5 != null && money5.compareTo(BigDecimal.ZERO) > 0) {
							money5 = money5.multiply(new BigDecimal(e.getPersonnelNum()))
									.divide(new BigDecimal(hasNumSum.get()), 4, BigDecimal.ROUND_HALF_UP);
							e.setCost5(e.getCost5().add(money5));
						}
						BigDecimal money6 = maps6.get(startDate);
						if (money6 != null && money6.compareTo(BigDecimal.ZERO) > 0) {
							money6 = money6.multiply(new BigDecimal(e.getPersonnelNum()))
									.divide(new BigDecimal(hasNumSum.get()), 4, BigDecimal.ROUND_HALF_UP);
							e.setCost6(e.getCost6().add(money6));
						}
						BigDecimal money7 = maps7.get(startDate);
						if (money7 != null && money7.compareTo(BigDecimal.ZERO) > 0) {
							money7 = money7.multiply(new BigDecimal(e.getPersonnelNum()))
									.divide(new BigDecimal(hasNumSum.get()), 2, BigDecimal.ROUND_HALF_UP);
							e.setCost7(e.getCost7().add(money7));
						}
						BigDecimal money8 = maps8.get(startDate);
						if (money8 != null && money8.compareTo(BigDecimal.ZERO) > 0) {
							money8 = money8.multiply(new BigDecimal(e.getPersonnelNum()))
									.divide(new BigDecimal(hasNumSum.get()), 4, BigDecimal.ROUND_HALF_UP);
							e.setCost8(e.getCost8().add(money8));
						}
						BigDecimal money9 = maps9.get(startDate);
						if (money9 != null && money9.compareTo(BigDecimal.ZERO) > 0) {
							money9 = money9.multiply(new BigDecimal(e.getPersonnelNum()))
									.divide(new BigDecimal(hasNumSum.get()), 4, BigDecimal.ROUND_HALF_UP);
							e.setCost9(e.getCost9().add(money9));
						}
						if (isMonth.get()) {
							startDate = DateUtils.addMonths(startDate, 1);
						} else {
							startDate = DateUtils.addDays(startDate, 1);
						}
					}
				}

			});

			projectCosts.forEach(System.out::println);
			writerExcle(projectCosts);
		}
	}

	@NotNull
	private Map<Date, BigDecimal> getDateCostMap(List<DateCost> dateCosts, AtomicBoolean isMonth) {
		Map<Date, BigDecimal> maps = new HashMap<>(50);
		dateCosts.forEach(e -> {
			Date costDate = DateUtils.parseDate(e.getCostDate());
			if (isMonth.get()) {
				costDate = DateUtils.parseDate(DateUtils.formatDate(costDate, "yyyy-MM"));
			}
			BigDecimal count = maps.get(costDate);
			if (count == null) {
				maps.put(costDate, e.getCost());
			} else {
				maps.put(costDate, e.getCost().add(count));
			}
		});
		maps.forEach((k, v) -> System.out.println(DateUtils.formatDate(k) + " = " + v));
		return maps;
	}

	private void writerExcle(List<ProjectCost> projectCosts) {
		EasyExcel.write(HandleExcel.outPath, ProjectCost.class).sheet("项目成本").doWrite(projectCosts);
	}

}
