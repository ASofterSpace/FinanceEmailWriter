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
import com.asofterspace.toolbox.io.IoUtils;
import com.asofterspace.toolbox.io.TextFile;
import com.asofterspace.toolbox.utils.Record;
import com.asofterspace.toolbox.utils.SortOrder;
import com.asofterspace.toolbox.utils.SortUtils;
import com.asofterspace.toolbox.utils.StrUtils;
import com.asofterspace.toolbox.Utils;
import com.asofterspace.toolbox.xlsx.XlsxFile;
import com.asofterspace.toolbox.xlsx.XlsxSheet;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
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
	public final static String VERSION_NUMBER = "0.0.0.6(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "6. June 2022 - 31. December 2022";


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

		IoUtils.cleanAllWorkDirs();

		Directory outputDirectory = new Directory("output");
		outputDirectory.clear();

		TextFile templateFile = new TextFile("input/template.txt");
		String templateText = templateFile.getContent();

		int amountOfPeople = 0;
		int amountOfNights = 0;
		int idealPayCounter = 0;
		int maxPayCounter = 0;
		int costCounterSum = 0;
		int costCounterHouse = 0;
		int costCounterFood = 0;
		int costCounterTransport = 0;
		int costCounterOther = 0;

		List<Person> people = new ArrayList<>();

		XlsxFile inputXlsx = new XlsxFile("input/input.xlsx");
		boolean usingXlsx = inputXlsx.exists();

		XlsxSheet sheetOutPayments = null;
		XlsxSheet sheetInData = null;
		XlsxSheet sheetInCalculation = null;
		XlsxSheet sheetInPayments = null;
		XlsxSheet sheetStatistics = null;

		if (usingXlsx) {
			// use XLSX file...
			List<XlsxSheet> sheets = inputXlsx.getSheets();
			for (XlsxSheet sheet : sheets) {
				switch(sheet.getTitle()) {

					case "OUT Payments":
						sheetOutPayments = sheet;
						break;

					case "IN Data":
						sheetInData = sheet;
						break;

					case "IN Calculation":
						sheetInCalculation = sheet;
						break;

					case "IN Payments":
						sheetInPayments = sheet;
						break;

					case "Statistics":
						sheetStatistics = sheet;
						break;
				}
			}

			if (sheetOutPayments == null) {
				System.out.println("Could not find OUT Payments sheet!");
				return;
			}
			if (sheetInData == null) {
				System.out.println("Could not find IN Data sheet!");
				return;
			}
			if (sheetInCalculation == null) {
				System.out.println("Could not find IN Calculation sheet!");
				return;
			}
			if (sheetInPayments == null) {
				System.out.println("Could not find IN Payments sheet!");
				return;
			}
			if (sheetStatistics == null) {
				System.out.println("Could not find Statistics sheet!");
				return;
			}

			Map<String, Record> persCol = sheetInData.getColContents("A");
			Set<String> keys = persCol.keySet();
			List<String> sortedKeys = SortUtils.sort(keys, SortOrder.NUMERICAL);
			for (String key : sortedKeys) {
				String rowNum = XlsxSheet.nameToRow(key);
				if (StrUtils.strToInt(rowNum) > 4) {
					Record nameRec = sheetInData.getCellContent("A" + rowNum);
					if (nameRec != null) {
						String name = nameRec.asString();
						if ("".equals(name)) {
							continue;
						}

						Person person = new Person(name);
						people.add(person);

						person.setContactMethod(sheetInData.getCellContent("B" + rowNum).asString());
						person.setIdealPay((int) Math.round(sheetInData.getCellContent("C" + rowNum).asDouble() * 100));
						person.setMaxPay((int) Math.round(sheetInData.getCellContent("D" + rowNum).asDouble() * 100));
						person.setTransInfo(sheetInData.getCellContent("E" + rowNum).asString());
						person.setTransCosts((int) Math.round(sheetInData.getCellContent("F" + rowNum).asDouble() * 100));
						person.setArrivalDate(sheetInData.getCellContent("G" + rowNum).asString());
						person.setLeaveDate(sheetInData.getCellContent("H" + rowNum).asString());
						person.setNights(sheetInData.getCellContent("I" + rowNum).asInteger());
						person.setHadExpense(0);
					}
				}
			}

			Map<String, Record> payCol = sheetOutPayments.getColContents("A");
			for (Map.Entry<String, Record> entry : payCol.entrySet()) {
				String key = entry.getKey();
				String rowNum = XlsxSheet.nameToRow(key);
				if (StrUtils.strToInt(rowNum) > 5) {
					Record amountRec = sheetOutPayments.getCellContent("A" + rowNum);
					if (amountRec != null) {
						Double amount = amountRec.asDouble();
						Record personRec = sheetOutPayments.getCellContent("B" + rowNum);
						String personStr = personRec.asString();
						Record catRec = sheetOutPayments.getCellContent("E" + rowNum);
						String catStr = null;
						if (catRec != null) {
							catStr = catRec.asString();
						}
						if (catStr == null) {
							catStr = "other";
						}
						if ("".equals(catStr)) {
							catStr = "o";
						}
						catStr = catStr.toLowerCase().trim().substring(0, 1);

						int amountInt = (int) Math.round(amount * 100);

						costCounterSum += amountInt;

						switch (catStr) {
							case "f":
								costCounterFood += amountInt;
								break;
							case "h":
								costCounterHouse += amountInt;
								break;
							case "t":
								costCounterTransport += amountInt;
								break;
							default:
								costCounterOther += amountInt;
								break;
						}

						boolean foundPerson = false;
						for (Person person : people) {
							if (personStr.equals(person.getName())) {
								person.setHadExpense(person.getHadExpense() + amountInt);
								foundPerson = true;
								break;
							}
						}

						System.out.println("payment row: " + rowNum + ", A: " + FinanceUtils.formatMoney(amountInt) + " €, P: " + personStr + ", C: " + catStr);

						if (!foundPerson) {
							System.out.println("No person with the name " + personStr + " found, cannot assign the expense!");
							return;
						}
					}
				}
			}

			// complain about people from IN Payments that are not on IN Data
			Map<String, Record> payInCol = sheetInPayments.getColContents("A");
			List<String> foundPeople = new ArrayList<>();
			for (Map.Entry<String, Record> entry : payInCol.entrySet()) {
				String key = entry.getKey();
				String rowNum = XlsxSheet.nameToRow(key);
				if (StrUtils.strToInt(rowNum) > 9) {
					Record val = entry.getValue();
					String name = val.asString();
					foundPeople.add(name);
					boolean foundThem = false;
					for (Person person : people) {
						if (person.getName().equals(name)) {
							foundThem = true;
							break;
						}
					}
					if (!foundThem) {
						System.out.println("The person " + name + " from IN Payments is not listed on IN Data!");
						return;
					}
				}
			}
			// add people to IN Payments who are on IN Data
			for (Person person : people) {
				if (!foundPeople.contains(person.getName())) {
					int rowNum = 10;
					while (sheetInPayments.getCellContent("A" + rowNum) != null) {
						rowNum++;
					}
					sheetInPayments.setCellContent("A" + rowNum, person.getName());
				}
			}

			int highestRowNum = sheetInCalculation.getHighestRowNum();
			sheetInCalculation.deleteCellBlock("A5", "J" + highestRowNum);

		} else {
			// fallback approach: use CSV input file...

			CsvFile inputCsv = new CsvFile("input/finance_sheet.csv");
			if (!inputCsv.exists()) {
				inputCsv = new CsvFile("input/input.csv");
			}
			inputCsv.setEntrySeparator(';');
			List<String> headline = inputCsv.getHeadLineInColumns();
			List<String> content = inputCsv.getContentLineInColumns();

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
				String name = getCell(content, COL_NAME);
				name = UniversalTextDecoder.decode(name).trim();

				if ("".equals(name.trim())) {
					break;
				}

				Person person = new Person(name);
				people.add(person);

				person.setContactMethod(getCell(content, COL_CONTACTMETH).trim());
				person.setIdealPay(FinanceUtils.parseMoney(getCell(content, COL_IDEAL)));
				person.setMaxPay(FinanceUtils.parseMoney(getCell(content, COL_MAX)));
				person.setHadExpense(FinanceUtils.parseMoney(getCell(content, COL_HAD_EXPENSE)));
				person.setTransInfo(getCell(content, COL_TRANSPORT_INFO));
				person.setTransCosts(FinanceUtils.parseMoney(getCell(content, COL_TRANSPORT_COSTS)));
				person.setNights(StrUtils.strToInt(getCell(content, COL_NIGHTS)));
				person.setArrivalDate(getCell(content, COL_ARRIVAL));
				person.setLeaveDate(getCell(content, COL_LEAVE));

				Integer agreedPay = FinanceUtils.parseMoney(getCell(content, COL_AGREED));
				costCounterSum += agreedPay;

				content = inputCsv.getContentLineInColumns();
			}
		}

		for (Person person : people) {
			amountOfPeople++;
			amountOfNights += person.getNights();
			idealPayCounter += person.getIdealPay();
			maxPayCounter += person.getMaxPay();
		}

		// calculate first round...
		double ratio1 = ((costCounterSum - maxPayCounter) * 1.0) / (idealPayCounter - maxPayCounter);
		ratio1 = ratio1 / 2;

		// adjust all ideal values based on first round, so that if payments are close to max, calc ideals are close
		// to ideal, and if payments are close to ideal, calc ideals are close to zero...
		int calcIdealPayCounter = 0;
		for (Person person : people) {
			person.setCalcIdealPay((int) Math.round(person.getIdealPay() * (1 - ratio1)));
			calcIdealPayCounter += person.getCalcIdealPay();
		}

		// calculate second round...
		double ratio2 = ((costCounterSum - maxPayCounter) * 1.0) / (calcIdealPayCounter - maxPayCounter);
		for (Person person : people) {
			person.setAgreedPay((int) Math.round((person.getCalcIdealPay() * ratio2) + (person.getMaxPay() * (1 - ratio2))));
		}

		System.out.println("Costs: " + FinanceUtils.formatMoney(costCounterSum) + " €");
		System.out.println("Ideal Sum: " + FinanceUtils.formatMoney(idealPayCounter) + " €");
		System.out.println("Max Sum: " + FinanceUtils.formatMoney(maxPayCounter) + " €");
		System.out.println("First Ratio: " + ratio1);
		System.out.println("Calculated Ideal Sum: " + FinanceUtils.formatMoney(calcIdealPayCounter) + " €");
		System.out.println("Second Ratio: " + ratio2);

		List<Integer> payList = new ArrayList<>();

		for (Person person : people) {
			payList.add(person.getAgreedPay());
		}

		Set<String> encounteredNames = new HashSet<>();
		int rollingRowNum = 5;

		for (Person person : people) {

			System.out.println(person.getName() + " (" + person.getContactMethod() + "), " +
				"ideal: " + FinanceUtils.formatMoney(person.getIdealPay()) + " €, " +
				"calc ideal: " + FinanceUtils.formatMoney(person.getCalcIdealPay()) + " €, " +
				"max: " + FinanceUtils.formatMoney(person.getMaxPay()) + " €, " +
				"payment: " + FinanceUtils.formatMoney(person.getAgreedPay()) + " €, " +
				"hadExpense: " + FinanceUtils.formatMoney(person.getHadExpense()) + " €, " +
				"(" + person.getNights() + ": " + person.getArrivalDate() + " - " + person.getLeaveDate() + ")");

			String outContent = templateText;
			outContent = StrUtils.replaceAll(outContent, CONTACT_METHOD, person.getContactMethod());
			outContent = StrUtils.replaceAll(outContent, NAME, person.getName());
			String idealPayStr = FinanceUtils.formatMoney(person.getIdealPay()) + " €";
			if (person.getNights() != 0) {
				idealPayStr += " (" + FinanceUtils.formatMoney(person.getIdealPay() / person.getNights()) + " € per night)";
			}
			outContent = StrUtils.replaceAll(outContent, IDEAL_PAY, idealPayStr);
			String maxPayStr = FinanceUtils.formatMoney(person.getMaxPay()) + " €";
			if (person.getNights() != 0) {
				maxPayStr += " (" + FinanceUtils.formatMoney(person.getMaxPay() / person.getNights()) + " € per night)";
			}
			outContent = StrUtils.replaceAll(outContent, MAX_PAY, maxPayStr);
			outContent = StrUtils.replaceAll(outContent, AGREED_PAY, FinanceUtils.formatMoney(person.getAgreedPay()) + " €");

			if (person.getHadExpense() == 0) {
				outContent = removeFromLineToLine(outContent, BEGIN_EXPENSE, END_EXPENSE);
			} else {
				outContent = removeLine(outContent, BEGIN_EXPENSE);
				outContent = removeLine(outContent, END_EXPENSE);
			}

			if (person.getTransCosts() == 0) {
				outContent = removeFromLineToLine(outContent, BEGIN_TRANS, END_TRANS);
			} else {
				outContent = removeLine(outContent, BEGIN_TRANS);
				outContent = removeLine(outContent, END_TRANS);
			}

			if ((person.getTransCosts() == 0) && (person.getHadExpense() == 0)) {
				outContent = removeFromLineToLine(outContent, BEGIN_EXPENSE_OR_TRANS, END_EXPENSE_OR_TRANS);
			} else {
				outContent = removeLine(outContent, BEGIN_EXPENSE_OR_TRANS);
				outContent = removeLine(outContent, END_EXPENSE_OR_TRANS);
			}

			outContent = StrUtils.replaceAll(outContent, HAD_EXPENSE, FinanceUtils.formatMoney(person.getHadExpense()) + " €");
			if (person.getActualTransaction() < 0) {
				outContent = StrUtils.replaceAll(outContent, ACTUAL_TRANSACTION, FinanceUtils.formatMoney(person.getActualTransaction()) +
					" € (as this amount is negative, I will actually send this much money to you, not the other way around ^^ - please let me know how you would like to get it.)");
				outContent = removeFromLineToLine(outContent, BEGIN_IF_POSITIVE, END_IF_POSITIVE);
			} else {
				outContent = StrUtils.replaceAll(outContent, ACTUAL_TRANSACTION, FinanceUtils.formatMoney(person.getActualTransaction()) + " €");
				outContent = removeLine(outContent, BEGIN_IF_POSITIVE);
				outContent = removeLine(outContent, END_IF_POSITIVE);
			}
			outContent = StrUtils.replaceAll(outContent, TRANSPORT_INFO, ""+person.getTransInfo());
			outContent = StrUtils.replaceAll(outContent, TRANSPORT_COSTS, FinanceUtils.formatMoney(person.getTransCosts()) + " €");
			outContent = StrUtils.replaceAll(outContent, NIGHTS, ""+person.getNights());
			outContent = StrUtils.replaceAll(outContent, ARRIVAL_DATE, person.getArrivalDate());
			outContent = StrUtils.replaceAll(outContent, LEAVE_DATE, person.getLeaveDate());

			String uniqueName = person.getName();
			while (encounteredNames.contains(uniqueName)) {
				uniqueName += " 2";
				System.out.println("Two people have the same name! Assigning payments will be hard!");
			}
			encounteredNames.add(uniqueName);

			TextFile outFile = new TextFile(outputDirectory, uniqueName + ".txt");
			outFile.saveContent(outContent);

			if (usingXlsx) {
				sheetInCalculation.setCellContent("A" + rollingRowNum, person.getName());
				sheetInCalculation.setCellContent("B" + rollingRowNum, person.getContactMethod());
				sheetInCalculation.setCellContent("C" + rollingRowNum, FinanceUtils.formatMoney(person.getIdealPay()) + " €");
				sheetInCalculation.setCellContent("D" + rollingRowNum, FinanceUtils.formatMoney(person.getCalcIdealPay()) + " €");
				sheetInCalculation.setCellContent("E" + rollingRowNum, FinanceUtils.formatMoney(person.getMaxPay()) + " €");
				sheetInCalculation.setCellContent("F" + rollingRowNum, FinanceUtils.formatMoney(person.getAgreedPay()) + " €");
				sheetInCalculation.setCellContent("G" + rollingRowNum, FinanceUtils.formatMoney(person.getHadExpense()) + " €");
				sheetInCalculation.setCellContent("H" + rollingRowNum, FinanceUtils.formatMoney(person.getTransCosts()) + " €");
				sheetInCalculation.setCellContent("I" + rollingRowNum, FinanceUtils.formatMoney(person.getActualTransaction()) + " €");
				sheetInCalculation.setCellContent("J" + rollingRowNum, person.getNights());
				sheetInCalculation.setCellContent("K" + rollingRowNum, FinanceUtils.formatMoney((int) Math.round((person.getAgreedPay() * 1.0) / person.getNights())) + " €");
				rollingRowNum++;
			}
		}

		int peopleAmountDiv10 = (int) Math.round(amountOfPeople / 10.0);
		List<Integer> sortedPayList = SortUtils.sortIntegers(payList);

		TextFile statsFile = new TextFile(outputDirectory, "_stats.txt");
		StringBuilder statsContent = new StringBuilder();
		statsContent.append("Stats for nerds:");
		statsContent.append("\r\n");
		statsContent.append(" ");
		statsContent.append("\r\n");
		statsContent.append("Amount to be paid overall: ");
		statsContent.append(FinanceUtils.formatMoney(costCounterSum) + " €");
		statsContent.append("\r\n");
		if (usingXlsx) {
			statsContent.append("Includes house costs: ");
			statsContent.append(FinanceUtils.formatMoney(costCounterHouse) + " €");
			statsContent.append("\r\n");
			statsContent.append("Includes food costs: ");
			statsContent.append(FinanceUtils.formatMoney(costCounterFood) + " €");
			statsContent.append("\r\n");
			if (costCounterTransport > 0) {
				statsContent.append("Includes transport costs: ");
				statsContent.append(FinanceUtils.formatMoney(costCounterTransport) + " €");
				statsContent.append("\r\n");
			}
			statsContent.append("Includes other costs: ");
			statsContent.append(FinanceUtils.formatMoney(costCounterOther) + " €");
			statsContent.append("\r\n");
		}
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
		statsContent.append(FinanceUtils.formatMoney(costCounterSum / amountOfPeople) + " €");
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
		statsContent.append(" (" + peopleAmountDiv10 + " " + (peopleAmountDiv10 == 1 ? "person pays" : "people pay") + " less than this amount)");
		statsContent.append("\r\n");
		statsContent.append("90th percentile: ");
		statsContent.append(FinanceUtils.formatMoney(
			(sortedPayList.get(amountOfPeople - peopleAmountDiv10) + sortedPayList.get(amountOfPeople - peopleAmountDiv10 - 1)) / 2
		) + " €");
		statsContent.append(" (" + peopleAmountDiv10 + " " + (peopleAmountDiv10 == 1 ? "person pays" : "people pay") + " more than this amount)");
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
		statsContent.append(FinanceUtils.formatMoney(costCounterSum / amountOfNights) + " €");
		statsContent.append("\r\n");
		statsFile.saveContent(statsContent.toString());

		if (usingXlsx) {
			sheetStatistics.setCellContent("A4", statsContent.toString());

			inputXlsx.saveTo(outputDirectory.getAbsoluteDirname() + "/_output.xlsx");
		}
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
