/**
 * Unlicensed code created by A Softer Space, 2022
 * www.asofterspace.com/licenses/unlicense.txt
 */
package com.asofterspace.financeEmailWriter;

import com.asofterspace.toolbox.accounting.FinanceUtils;
import com.asofterspace.toolbox.utils.Pair;
import com.asofterspace.toolbox.utils.StrUtils;

import java.util.ArrayList;
import java.util.List;


public class Person {

	private String displayName;
	private String name;
	private String contactMethod;
	private Integer idealPay;
	private Integer calcIdealPay;
	private Integer maxPay;
	private Integer agreedPay;
	private Integer hadExpense;
	private String transInfo;
	private Integer transCosts;
	private Integer nights;
	private String arrivalDate;
	private String leaveDate;

	private List<Pair<Integer, String>> expenses;
	private List<Pair<Integer, String>> transports;


	public Person(String name) {
		this.displayName = name;
		this.name = sanitizeName(name);
		this.expenses = new ArrayList<>();
		this.hadExpense = 0;
		this.transports = new ArrayList<>();
		this.transCosts = 0;
	}

	public static String sanitizeName(String name) {
		name = StrUtils.removeTrailingPronounsFromName(name);
		name = StrUtils.replaceAll(name, "/", "");
		name = StrUtils.replaceAll(name, ".", "");
		name = StrUtils.replaceAll(name, "  ", " ");
		name = name.trim();
		return name;
	}

	public String getName() {
		return name;
	}

	public String getDisplayName() {
		return displayName;
	}

	public String getContactMethod() {
		return contactMethod;
	}

	public void setContactMethod(String contactMethod) {
		this.contactMethod = contactMethod;
	}

	public Integer getIdealPay() {
		return idealPay;
	}

	public void setIdealPay(Integer idealPay) {
		this.idealPay = idealPay;
	}

	public Integer getCalcIdealPay() {
		return calcIdealPay;
	}

	public void setCalcIdealPay(Integer calcIdealPay) {
		this.calcIdealPay = calcIdealPay;
	}

	public Integer getMaxPay() {
		return maxPay;
	}

	public void setMaxPay(Integer maxPay) {
		this.maxPay = maxPay;
	}

	public Integer getAgreedPay() {
		return agreedPay;
	}

	public void setAgreedPay(Integer agreedPay) {
		this.agreedPay = agreedPay;
	}

	public Integer getHadExpense() {
		return hadExpense;
	}

	public void setHadExpense(Integer hadExpense) {
		this.hadExpense = hadExpense;
	}

	public void addExpense(int costs, String text) {
		expenses.add(new Pair<>(costs, text));
		this.hadExpense += costs;
	}

	public String getExpenseText() {
		String result = "";

		if (expenses.size() < 1) {
			if (getHadExpense() == 0) {
				return "(none)";
			}
			return FinanceUtils.formatMoney(getHadExpense()) + " €";
		}

		String sep = "";

		for (Pair<Integer, String> cost : expenses) {
			result += sep + FinanceUtils.formatMoney(cost.getKey()) + " € for " + cost.getValue();
			sep = "\r\n";
		}

		// only show a total if there is more than one row
		if (expenses.size() > 1) {
			result += sep + "So in total: " + FinanceUtils.formatMoney(getHadExpense()) + " € in expenses.";
		}

		return result;
	}

	public String getTransInfo() {
		return transInfo;
	}

	public void setTransInfo(String transInfo) {
		this.transInfo = transInfo;
	}

	public Integer getTransCosts() {
		return transCosts;
	}

	public void setTransCosts(Integer transCosts) {
		this.transCosts = transCosts;
	}

	public void addTransport(int costs, String text) {
		transports.add(new Pair<>(costs, text));
		this.transCosts += costs;
	}

	public String getTransportText() {
		String result = "";

		if (transports.size() < 1) {
			return FinanceUtils.formatMoney(getTransCosts()) + " € for " + getTransInfo();
		}

		String sep = "";

		for (Pair<Integer, String> cost : transports) {
			result += sep + FinanceUtils.formatMoney(cost.getKey()) + " € for " + cost.getValue();
			sep = "\r\n";
		}

		// only show a total if there is more than one row
		if (transports.size() > 1) {
			result += sep + "So in total: " + FinanceUtils.formatMoney(getTransCosts()) + " € in transport costs.";
		}

		return result;
	}

	public Integer getNights() {
		return nights;
	}

	public void setNights(Integer nights) {
		if (nights == null) {
			this.nights = 0;
		} else {
			this.nights = nights;
		}
	}

	public String getArrivalDate() {
		return arrivalDate;
	}

	public void setArrivalDate(String arrivalDate) {
		this.arrivalDate = arrivalDate;
	}

	public String getLeaveDate() {
		return leaveDate;
	}

	public void setLeaveDate(String leaveDate) {
		this.leaveDate = leaveDate;
	}

	public Integer getActualTransaction() {
		return agreedPay - hadExpense - transCosts;
	}

}
