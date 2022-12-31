/**
 * Unlicensed code created by A Softer Space, 2022
 * www.asofterspace.com/licenses/unlicense.txt
 */
package com.asofterspace.financeEmailWriter;

import com.asofterspace.toolbox.utils.StrUtils;


public class Person {

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


	public Person(String name) {
		this.name = StrUtils.removeTrailingPronounsFromName(name);
		this.name = StrUtils.replaceAll(this.name, "/", "");
		this.name = StrUtils.replaceAll(this.name, "  ", " ");
		this.name = this.name.trim();
	}

	public String getName() {
		return name;
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

	public Integer getNights() {
		return nights;
	}

	public void setNights(Integer nights) {
		this.nights = nights;
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
