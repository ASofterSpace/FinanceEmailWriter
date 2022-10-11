/**
 * Unlicensed code created by A Softer Space, 2022
 * www.asofterspace.com/licenses/unlicense.txt
 */
package com.asofterspace.financeEmailWriter;

import com.asofterspace.toolbox.accounting.FinanceUtils;
import com.asofterspace.toolbox.coders.UniversalTextDecoder;
import com.asofterspace.toolbox.io.CsvFile;
import com.asofterspace.toolbox.io.Directory;
import com.asofterspace.toolbox.io.File;
import com.asofterspace.toolbox.io.TextFile;
import com.asofterspace.toolbox.utils.SortUtils;
import com.asofterspace.toolbox.utils.StrUtils;
import com.asofterspace.toolbox.Utils;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;


public class FinanceEmailWriter {

	// fields in email texts
	private static final String CONTACT_METHOD = "(CONTACT_METHOD)";
	private static final String NAME = "(NAME)";
	private static final String IDEAL_PAY = "(IDEAL_PAY)";
	private static final String MAX_PAY = "(MAX_PAY)";
	private static final String AGREED_PAY = "(AGREED_PAY)";
	private static final String BEGIN_EXPENSE = "(BEGIN_EXPENSE)";
	private static final String END_EXPENSE = "(END_EXPENSE)";
	private static final String BEGIN_TRANS = "(BEGIN_TRANSPORT_COSTS)";
	private static final String END_TRANS = "(END_TRANSPORT_COSTS)";
	private static final String BEGIN_EXPENSE_OR_TRANS = "(BEGIN_EXPENSE_OR_TRANS)";
	private static final String END_EXPENSE_OR_TRANS = "(END_EXPENSE_OR_TRANS)";
	private static final String HAD_EXPENSE = "(HAD_EXPENSE)";
	private static final String TRANSPORT_INFO = "(TRANSPORT_INFO)";
	private static final String TRANSPORT_COSTS = "(TRANSPORT_COSTS)";
	private static final String BEGIN_IF_POSITIVE = "(BEGIN_IF_POSITIVE)";
	private static final String END_IF_POSITIVE = "(END_IF_POSITIVE)";
	private static final String ACTUAL_TRANSACTION = "(ACTUAL_TRANSACTION)";
	private static final String NIGHTS = "(NIGHTS)";
	private static final String ARRIVAL_DATE = "(ARRIVAL_DATE)";
	private static final String LEAVE_DATE = "(LEAVE_DATE)";

	public final static String PROGRAM_TITLE = "FinanceEmailWriter";
	public final static String VERSION_NUMBER = "0.0.0.5(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "6. June 2022 - 11. October 2022";


	public static void main(String[] args) throws Exception {

		// let the Utils know in what program it is being used
		Utils.setProgramTitle(PROGRAM_TITLE);
		Utils.setVersionNumber(VERSION_NUMBER);
		Utils.setVersionDate(VERSION_DATE);

		if (args.length > 0) {
			if (args[0].equals("--version")) {
				System.out.println(Utils.getFullProgramIdentifierWithDate());
				return;
			}

			if (args[0].equals("--version_for_zip")) {
				System.out.println("version " + Utils.getVersionNumber());
				return;
			}
		}

		Directory outputDir = new Directory("output");
		outputDir.clear();

		Set<String> encounteredNames = new HashSet<>();

		TextFile templateFile = new TextFile("input/template.txt");
		String templateText = templateFile.getContent();

		CsvFile inputFile = new CsvFile("input/finance_sheet.csv");
		inputFile.setEntrySeparator(';');
		List<String> headline = inputFile.getHeadLineInColumns();
		List<String> content = inputFile.getContentLineInColumns();

		int amountOfPeople = 0;
		int amountOfNights = 0;
		int idealPayCounter = 0;
		int maxPayCounter = 0;
		int avgPayCounter = 0;
		List<Integer> payList = new ArrayList<>();

		// fields in input CSV files
		Integer COL_CONTACTMETH = getColNum(headline, "email");
		Integer COL_NAME = getColNum(headline, "human");
		Integer COL_IDEAL = getColNum(headline, "ideal budget");
		Integer COL_MAX = getColNum(headline, "maximum budget");
		Integer COL_AGREED = getColNum(headline, "agreed payment");
		Integer COL_HAD_EXPENSE = getColNum(headline, "had expense");
		Integer COL_TRANSPORT_INFO = getColNum(headline, "transport info");
		Integer COL_TRANSPORT_COSTS = getColNum(headline, "transport costs");
		Integer COL_ACTUAL = getColNum(headline, "actual transaction");
		Integer COL_NIGHTS = getColNum(headline, "nights");
		Integer COL_ARRIVAL = getColNum(headline, "arrival");
		Integer COL_LEAVE = getColNum(headline, "departure");

		while (content != null) {
			String contactMethod = getCell(content, COL_CONTACTMETH).trim();
			String name = getCell(content, COL_NAME);
			name = UniversalTextDecoder.decode(name).trim();
			name = StrUtils.removeTrailingPronounsFromName(name);
			Integer idealPay = FinanceUtils.parseMoney(getCell(content, COL_IDEAL));
			Integer maxPay = FinanceUtils.parseMoney(getCell(content, COL_MAX));
			Integer agreedPay = FinanceUtils.parseMoney(getCell(content, COL_AGREED));
			Integer hadExpense = FinanceUtils.parseMoney(getCell(content, COL_HAD_EXPENSE));
			String transInfo = getCell(content, COL_TRANSPORT_INFO);
			Integer transCosts = FinanceUtils.parseMoney(getCell(content, COL_TRANSPORT_COSTS));
			Integer actualTransaction = FinanceUtils.parseMoney(getCell(content, COL_ACTUAL));

			Integer nights = StrUtils.strToInt(getCell(content, COL_NIGHTS));
			String arrivalDate = getCell(content, COL_ARRIVAL);
			String leaveDate = getCell(content, COL_LEAVE);

			if ("".equals(name.trim())) {
				break;
			}

			amountOfPeople++;
			amountOfNights += nights;
			idealPayCounter += idealPay;
			maxPayCounter += maxPay;
			avgPayCounter += agreedPay;
			payList.add(agreedPay);

			String uniqueName = name;
			uniqueName = StrUtils.replaceAll(uniqueName, "/", "");
			uniqueName = StrUtils.replaceAll(uniqueName, "  ", " ");
			uniqueName = uniqueName.trim();

			while (encounteredNames.contains(uniqueName)) {
				uniqueName += " 2";
			}
			encounteredNames.add(uniqueName);

			System.out.println(name + " (" + contactMethod + "), " +
				"ideal: " + FinanceUtils.formatMoney(idealPay) + " €, " +
				"max: " + FinanceUtils.formatMoney(maxPay) + " €, " +
				"hadExpense: " + FinanceUtils.formatMoney(hadExpense) + " €, " +
				"(" + nights + ": " + arrivalDate + " - " + leaveDate + ")");

			int calcTransaction = agreedPay - hadExpense - transCosts;
			if (calcTransaction != actualTransaction) {
				System.out.println("Calculated different actual transaction for " + name + " than the input file has! " +
					"(calculated: " + FinanceUtils.formatMoney(calcTransaction) + " €, " +
					"actual: " + FinanceUtils.formatMoney(actualTransaction) + " €)");
			}

			String outContent = templateText;
			outContent = StrUtils.replaceAll(outContent, CONTACT_METHOD, contactMethod);
			outContent = StrUtils.replaceAll(outContent, NAME, name);
			String idealPayStr = FinanceUtils.formatMoney(idealPay) + " €";
			if (nights != 0) {
				idealPayStr += " (" + FinanceUtils.formatMoney(idealPay / nights) + " € per night)";
			}
			outContent = StrUtils.replaceAll(outContent, IDEAL_PAY, idealPayStr);
			String maxPayStr = FinanceUtils.formatMoney(maxPay) + " €";
			if (nights != 0) {
				maxPayStr += " (" + FinanceUtils.formatMoney(maxPay / nights) + " € per night)";
			}
			outContent = StrUtils.replaceAll(outContent, MAX_PAY, maxPayStr);
			outContent = StrUtils.replaceAll(outContent, AGREED_PAY, FinanceUtils.formatMoney(agreedPay) + " €");

			if (hadExpense == 0) {
				outContent = removeFromLineToLine(outContent, BEGIN_EXPENSE, END_EXPENSE);
			} else {
				outContent = removeLine(outContent, BEGIN_EXPENSE);
				outContent = removeLine(outContent, END_EXPENSE);
			}

			if (transCosts == 0) {
				outContent = removeFromLineToLine(outContent, BEGIN_TRANS, END_TRANS);
			} else {
				outContent = removeLine(outContent, BEGIN_TRANS);
				outContent = removeLine(outContent, END_TRANS);
			}

			if ((transCosts == 0) && (hadExpense == 0)) {
				outContent = removeFromLineToLine(outContent, BEGIN_EXPENSE_OR_TRANS, END_EXPENSE_OR_TRANS);
			} else {
				outContent = removeLine(outContent, BEGIN_EXPENSE_OR_TRANS);
				outContent = removeLine(outContent, END_EXPENSE_OR_TRANS);
			}

			outContent = StrUtils.replaceAll(outContent, HAD_EXPENSE, FinanceUtils.formatMoney(hadExpense) + " €");
			if (actualTransaction < 0) {
				outContent = StrUtils.replaceAll(outContent, ACTUAL_TRANSACTION, FinanceUtils.formatMoney(actualTransaction) +
					" € (as this amount is negative, I will actually send this much money to you, not the other way around ^^ - please let me know how you would like to get it.)");
				outContent = removeFromLineToLine(outContent, BEGIN_IF_POSITIVE, END_IF_POSITIVE);
			} else {
				outContent = StrUtils.replaceAll(outContent, ACTUAL_TRANSACTION, FinanceUtils.formatMoney(actualTransaction) + " €");
				outContent = removeLine(outContent, BEGIN_IF_POSITIVE);
				outContent = removeLine(outContent, END_IF_POSITIVE);
			}
			outContent = StrUtils.replaceAll(outContent, TRANSPORT_INFO, ""+transInfo);
			outContent = StrUtils.replaceAll(outContent, TRANSPORT_COSTS, FinanceUtils.formatMoney(transCosts) + " €");
			outContent = StrUtils.replaceAll(outContent, NIGHTS, ""+nights);
			outContent = StrUtils.replaceAll(outContent, ARRIVAL_DATE, arrivalDate);
			outContent = StrUtils.replaceAll(outContent, LEAVE_DATE, leaveDate);

			TextFile outFile = new TextFile(outputDir, uniqueName + ".txt");
			outFile.saveContent(outContent);

			content = inputFile.getContentLineInColumns();
		}

		int peopleAmountDiv10 = (int) Math.round(amountOfPeople / 10.0);
		List<Integer> sortedPayList = SortUtils.sortIntegers(payList);

		TextFile statsFile = new TextFile(outputDir, "_stats.txt");
		StringBuilder statsContent = new StringBuilder();
		statsContent.append("Stats for nerds:");
		statsContent.append("\r\n");
		statsContent.append(" ");
		statsContent.append("\r\n");
		statsContent.append("Amount to be paid overall: ");
		statsContent.append(FinanceUtils.formatMoney(avgPayCounter) + " €");
		statsContent.append("\r\n");
		statsContent.append("Amount of people: " + amountOfPeople);
		statsContent.append("\r\n");
		statsContent.append(" ");
		statsContent.append("\r\n");
		statsContent.append("Average ideal payment: ");
		statsContent.append(FinanceUtils.formatMoney(idealPayCounter / amountOfPeople) + " €");
		statsContent.append("\r\n");
		statsContent.append("Average maximum payment: ");
		statsContent.append(FinanceUtils.formatMoney(maxPayCounter / amountOfPeople) + " €");
		statsContent.append("\r\n");
		statsContent.append("Average payment: ");
		statsContent.append(FinanceUtils.formatMoney(avgPayCounter / amountOfPeople) + " €");
		statsContent.append("\r\n");
		statsContent.append("Median payment: ");
		if (amountOfPeople % 2 == 0) {
			statsContent.append(FinanceUtils.formatMoney(
				(sortedPayList.get(((int) Math.floor(amountOfPeople / 2)) - 1) +
				sortedPayList.get((int) Math.floor(amountOfPeople / 2))) / 2
			) + " €");
		} else {
			statsContent.append(FinanceUtils.formatMoney(
				sortedPayList.get((int) Math.floor(amountOfPeople / 2))
			) + " €");
		}
		statsContent.append("\r\n");
		statsContent.append("10th percentile: ");
		statsContent.append(FinanceUtils.formatMoney(
			(sortedPayList.get(peopleAmountDiv10) + sortedPayList.get(peopleAmountDiv10 + 1)) / 2
		) + " €");
		statsContent.append(" (" + peopleAmountDiv10 + " people pay less than this amount)");
		statsContent.append("\r\n");
		statsContent.append("90th percentile: ");
		statsContent.append(FinanceUtils.formatMoney(
			(sortedPayList.get(amountOfPeople - peopleAmountDiv10) + sortedPayList.get(amountOfPeople - peopleAmountDiv10 - 1)) / 2
		) + " €");
		statsContent.append(" (" + peopleAmountDiv10 + " people pay more than this amount)");
		statsContent.append("\r\n");
		statsContent.append(" ");
		statsContent.append("\r\n");
		statsContent.append("Average amount of nights that people are at the event: ");
		statsContent.append(StrUtils.doubleToStr((1.0 * amountOfNights) / amountOfPeople, 2));
		statsContent.append("\r\n");
		statsContent.append("Average ideal payment per night: ");
		statsContent.append(FinanceUtils.formatMoney(idealPayCounter / amountOfNights) + " €");
		statsContent.append("\r\n");
		statsContent.append("Average maximum payment per night: ");
		statsContent.append(FinanceUtils.formatMoney(maxPayCounter / amountOfNights) + " €");
		statsContent.append("\r\n");
		statsContent.append("Average payment per night: ");
		statsContent.append(FinanceUtils.formatMoney(avgPayCounter / amountOfNights) + " €");
		statsContent.append("\r\n");
		statsFile.saveContent(statsContent.toString());
	}

	private static Integer getColNum(List<String> headline, String columnStr) {

		if (headline == null) {
			return null;
		}

		String lowNeedle = columnStr.toLowerCase().trim();

		for (int i = 0; i < headline.size(); i++) {
			String headlineStr = headline.get(i).toLowerCase().trim();

			if (lowNeedle.equals(headlineStr)) {
				return i;
			}
		}
		return null;
	}

	private static String getCell(List<String> content, Integer column) {
		if (column == null) {
			return "";
		}
		if (content == null) {
			return "";
		}
		String result = content.get(column);
		if (result == null) {
			return "";
		}
		return result;
	}

	private static String removeLine(String outContent, String line) {
		outContent = StrUtils.replaceAll(outContent, line + "\r\n", "");
		outContent = StrUtils.replaceAll(outContent, line + "\n", "");
		return outContent;
	}

	private static String removeFromLineToLine(String outContent, String fromLine, String toLine) {

		int start = outContent.indexOf(fromLine + "\r\n");
		int end = outContent.indexOf(toLine + "\r\n");
		if ((start > -1) && (end > -1)) {
			outContent = outContent.substring(0, start) + outContent.substring(end + toLine.length() + 2);
		}

		start = outContent.indexOf(fromLine + "\n");
		end = outContent.indexOf(toLine + "\n");
		if ((start > -1) && (end > -1)) {
			outContent = outContent.substring(0, start) + outContent.substring(end + toLine.length() + 1);
		}

		return outContent;
	}
}
