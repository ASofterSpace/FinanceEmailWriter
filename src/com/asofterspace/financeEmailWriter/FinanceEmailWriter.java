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
import com.asofterspace.toolbox.utils.MathUtils;
import com.asofterspace.toolbox.utils.Record;
import com.asofterspace.toolbox.utils.SortOrder;
import com.asofterspace.toolbox.utils.SortUtils;
import com.asofterspace.toolbox.utils.StrUtils;
import com.asofterspace.toolbox.Utils;
import com.asofterspace.toolbox.xlsx.XlsxFile;
import com.asofterspace.toolbox.xlsx.XlsxSheet;

import java.util.ArrayList;
import java.util.HashMap;
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
	private static final String BEGIN_IF_NEGATIVE = "(BEGIN_IF_NEGATIVE)";
	private static final String END_IF_NEGATIVE = "(END_IF_NEGATIVE)";
	private static final String ACTUAL_TRANSACTION = "(ACTUAL_TRANSACTION)";
	private static final String NIGHTS = "(NIGHTS)";
	private static final String ARRIVAL_DATE = "(ARRIVAL_DATE)";
	private static final String LEAVE_DATE = "(LEAVE_DATE)";
	private static final String BEGIN_IF_NIGHTS_ZERO = "(BEGIN_IF_NIGHTS_ZERO)";
	private static final String END_IF_NIGHTS_ZERO = "(END_IF_NIGHTS_ZERO)";
	private static final String BEGIN_IF_NIGHTS_NON_ZERO = "(BEGIN_IF_NIGHTS_NON_ZERO)";
	private static final String END_IF_NIGHTS_NON_ZERO = "(END_IF_NIGHTS_NON_ZERO)";
	private static final int HEAD_LINE_AMOUNT_IN_PAYMENTS = 9;

	public final static String PROGRAM_TITLE = "FinanceEmailWriter";
	public final static String VERSION_NUMBER = "0.0.1.2(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "6. June 2022 - 30. April 2023";


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
		templateFile.copyToDisk(new File(outputDirectory, "_template.txt"));

		int amountOfPeople = 0;
		int amountOfNights = 0;
		int idealPayCounter = 0;
		int maxPayCounter = 0;
		int costCounterSum = 0;
		int costCounterHouse = 0;
		int costCounterFood = 0;
		int costCounterTransport = 0;
		int costCounterCredit = 0;
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

						person.setContactMethod(sheetInData.getCellContentString("B" + rowNum));
						person.setIdealPay((int) Math.round(sheetInData.getCellContentDoubleNonNull("C" + rowNum) * 100));
						person.setMaxPay((int) Math.round(sheetInData.getCellContentDoubleNonNull("D" + rowNum) * 100));
						person.setArrivalDate(sheetInData.getCellContentString("E" + rowNum));
						person.setLeaveDate(sheetInData.getCellContentString("F" + rowNum));
						person.setNights(sheetInData.getCellContentInteger("G" + rowNum));
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
						String personStr = Person.sanitizeName(personRec.asString());
						String expenseText = sheetOutPayments.getCellContentStringNonNull("D" + rowNum);
						String catStr = sheetOutPayments.getCellContentStringNonNull("E" + rowNum);
						if ("".equals(catStr)) {
							catStr = "o";
						}
						catStr = catStr.toLowerCase().trim().substring(0, 1);

						int amountInt = (int) Math.round(amount * 100);

						costCounterSum += amountInt;

						switch (catStr) {
							// Food
							case "f":
								costCounterFood += amountInt;
								break;
							// House
							case "h":
								costCounterHouse += amountInt;
								break;
							// Transport
							case "t":
								costCounterTransport += amountInt;
								break;
							// Credit
							case "c":
								costCounterCredit += amountInt;
								break;
							// Other
							default:
								costCounterOther += amountInt;
								break;
						}

						boolean foundPerson = false;
						for (Person person : people) {
							if (personStr.equals(person.getName())) {
								if ("t".equals(catStr)) {
									person.addTransport(amountInt, expenseText);
								} else {
									person.addExpense(amountInt, expenseText);
								}
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
					if (val == null) {
						continue;
					}
					String name = val.asString();
					name = Person.sanitizeName(name);
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

		// calculate new ideal values by calculating the average ratio of ideal to max,
		// and then calculating new ideal values for each person based on that ratio and their max value,
		// but with more weight for the ideal value they originally gave
		// (so max values remain, but people who put very low ideals will get them slightly raised, and people who
		// put very high ideals will get them slightly lowered, to have overall a fairer distribution in which
		// everyone benefits from the community more equally than they otherwise would)
		double overallIdealToMaxRatio = (idealPayCounter * 1.0) / maxPayCounter;

		for (Person person : people) {
			// special case: if orig ideal == orig max, then the person really wants to pay EXACTLY
			// that amount, so let their calc ideal also be equal to that value so that they really
			// to get this amount, no matter what!
			if (MathUtils.equals(person.getIdealPay(), person.getMaxPay())) {
				person.setCalcIdealPay(person.getIdealPay());
				continue;
			}

			double fullyCalculatedPay = overallIdealToMaxRatio * person.getMaxPay();
			double avgCalcAndOrigIdealPay = ((2 * person.getIdealPay()) + fullyCalculatedPay) / 3;
			person.setCalcIdealPay((int) Math.round(avgCalcAndOrigIdealPay));
		}

		boolean repeatCalculation = true;
		int calcIdealPayCounter = 0;
		double interpolationFactor = 0;
		int roundCounter = 0;

		// calculate actual interpolation, based on calculated ideal and original max value...
		while (repeatCalculation) {

			roundCounter++;

			System.out.println("Running interpolation round #" + roundCounter + "...");

			repeatCalculation = false;

			calcIdealPayCounter = 0;
			for (Person person : people) {
				calcIdealPayCounter += person.getCalcIdealPay();
			}

			interpolationFactor = ((costCounterSum - maxPayCounter) * 1.0) / (calcIdealPayCounter - maxPayCounter);
			for (Person person : people) {
				person.setAgreedPay((int) Math.round((person.getCalcIdealPay() * interpolationFactor) +
					(person.getMaxPay() * (1 - interpolationFactor))));
			}

			// ... but do this again and again, as long as the outcoming values are below 10% above the ideal pay
			// (unless they are for everyone, in which case this will stop running once all the calc ideals
			// are at the orig ideals)
			for (Person person : people) {
				int tenPercAboveIdeal = person.getIdealPay() + ((person.getMaxPay() - person.getIdealPay()) / 10);
				if (person.getAgreedPay() < tenPercAboveIdeal) {
					if (person.getCalcIdealPay() < person.getIdealPay()) {
						int upStepSize = (person.getMaxPay() - person.getCalcIdealPay()) / 20;
						// ensure we don't just loop forever
						if (upStepSize < 1) {
							upStepSize = 1;
						}
						person.setCalcIdealPay(Math.min(
							person.getCalcIdealPay() + upStepSize,
							person.getIdealPay()
						));
						repeatCalculation = true;
					}
				}
			}
		}

		System.out.println("Costs: " + FinanceUtils.formatMoney(costCounterSum) + " €");
		System.out.println("Orig Ideal Sum: " + FinanceUtils.formatMoney(idealPayCounter) + " €");
		System.out.println("Orig Max Sum: " + FinanceUtils.formatMoney(maxPayCounter) + " €");
		System.out.println("Overall Orig Ideal to Orig Max Ratio: " + overallIdealToMaxRatio);
		System.out.println("Calculated Ideal Sum: " + FinanceUtils.formatMoney(calcIdealPayCounter) + " €");
		System.out.println("Interpolation Factor between Calc Ideal and Orig Max: " + interpolationFactor);

		List<Integer> payList = new ArrayList<>();

		for (Person person : people) {
			payList.add(person.getAgreedPay());
		}

		Set<String> encounteredNames = new HashSet<>();
		int rollingRowNum = 5;

		if (usingXlsx) {
			// sort people on IN Payments alphabetically - but only if the rest of the rows is empty!
			// soooo check that...
			boolean isEmpty = true;
			Map<String, List<String>> rows = new HashMap<>();

			Map<String, Record> payInCol = sheetInPayments.getColContents("A");
			for (Map.Entry<String, Record> entry : payInCol.entrySet()) {
				String key = entry.getKey();
				String rowNum = XlsxSheet.nameToRow(key);
				if (StrUtils.strToInt(rowNum) > HEAD_LINE_AMOUNT_IN_PAYMENTS) {
					Record val = entry.getValue();
					if (val == null) {
						continue;
					}
					String name = val.asString();
					List<String> cellList = new ArrayList<>();
					String bContentStr = sheetInPayments.getCellContentString("B" + rowNum);
					cellList.add(bContentStr);
					if (!((bContentStr == null) || "".equals(bContentStr))) {
						isEmpty = false;
					}
					String cContentStr = sheetInPayments.getCellContentString("C" + rowNum);
					cellList.add(cContentStr);
					if (!((cContentStr == null) || "".equals(cContentStr))) {
						isEmpty = false;
					}
					cellList.add(sheetInPayments.getCellContentString("D" + rowNum));
					cellList.add(sheetInPayments.getCellContentString("F" + rowNum));
					cellList.add(sheetInPayments.getCellContentString("G" + rowNum));
					rows.put(name, cellList);
				}
			}

			if (isEmpty) {
				Set<String> names = rows.keySet();
				List<String> sortedNames = SortUtils.sort(names, SortOrder.ALPHABETICAL_IGNORE_UMLAUTS);
				int cur = HEAD_LINE_AMOUNT_IN_PAYMENTS + 1;
				for (String sortedName : sortedNames) {
					for (Map.Entry<String, List<String>> row : rows.entrySet()) {
						if (sortedName.equals(row.getKey())) {
							List<String> cellList = row.getValue();
							sheetInPayments.setCellContent("A" + cur, row.getKey());
							sheetInPayments.setCellContent("B" + cur, cellList.get(0));
							sheetInPayments.setCellContent("C" + cur, cellList.get(1));
							sheetInPayments.setCellContent("D" + cur, cellList.get(2));
							sheetInPayments.setCellContent("F" + cur, cellList.get(3));
							sheetInPayments.setCellContent("G" + cur, cellList.get(4));
							cur++;
							continue;
						}
					}
				}
				System.out.println("Sorting people alphabetically on IN Payments sheet...");
			} else {
				System.out.println("Not sorting people alphabetically on IN Payments sheet, " +
					"as there is already some info there!");
			}
		}

		int amountOfPeopleZeroNights = 0;

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
			outContent = StrUtils.replaceAll(outContent, NAME, person.getDisplayName());

			if (person.getNights() == 0) {
				amountOfPeopleZeroNights++;
				outContent = StrUtils.replaceAll(outContent, BEGIN_IF_NIGHTS_ZERO, "");
				outContent = StrUtils.replaceAll(outContent, END_IF_NIGHTS_ZERO, "");
				outContent = removeFromTo(outContent, BEGIN_IF_NIGHTS_NON_ZERO, END_IF_NIGHTS_NON_ZERO);
			} else {
				outContent = StrUtils.replaceAll(outContent, BEGIN_IF_NIGHTS_NON_ZERO, "");
				outContent = StrUtils.replaceAll(outContent, END_IF_NIGHTS_NON_ZERO, "");
				outContent = removeFromTo(outContent, BEGIN_IF_NIGHTS_ZERO, END_IF_NIGHTS_ZERO);
			}

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

			outContent = StrUtils.replaceAll(outContent, HAD_EXPENSE, person.getExpenseText());
			// no longer necessary, as this is just part of the transport costs now
			// outContent = StrUtils.replaceAll(outContent, TRANSPORT_INFO, ""+person.getTransInfo());
			outContent = StrUtils.replaceAll(outContent, TRANSPORT_COSTS,person.getTransportText());
			outContent = StrUtils.replaceAll(outContent, ACTUAL_TRANSACTION, FinanceUtils.formatMoney(person.getActualTransaction()) + " €");
			if (person.getActualTransaction() < 0) {
				outContent = removeFromLineToLine(outContent, BEGIN_IF_POSITIVE, END_IF_POSITIVE);
				outContent = removeLine(outContent, BEGIN_IF_NEGATIVE);
				outContent = removeLine(outContent, END_IF_NEGATIVE);
			} else {
				outContent = removeLine(outContent, BEGIN_IF_POSITIVE);
				outContent = removeLine(outContent, END_IF_POSITIVE);
				outContent = removeFromLineToLine(outContent, BEGIN_IF_NEGATIVE, END_IF_NEGATIVE);
			}
			outContent = StrUtils.replaceAll(outContent, NIGHTS, ""+person.getNights());
			outContent = StrUtils.replaceAll(outContent, ARRIVAL_DATE, person.getArrivalDate());
			outContent = StrUtils.replaceAll(outContent, LEAVE_DATE, person.getLeaveDate());

			String uniqueName = person.getName();
			while (encounteredNames.contains(uniqueName)) {
				uniqueName += " 2";
				System.out.println("Two people have the same name! Assigning payments will be hard!");
			}
			encounteredNames.add(uniqueName);

			// sanitize whitespaces just in case
			outContent = StrUtils.replaceAll(outContent, "  ", " ");

			TextFile outFile = new TextFile(outputDirectory, uniqueName + ".txt");
			outFile.saveContent(outContent);

			if (usingXlsx) {
				sheetInCalculation.setCellContent("A" + rollingRowNum, person.getName());
				sheetInCalculation.setCellContent("B" + rollingRowNum, person.getContactMethod());
				sheetInCalculation.setCellContent("C" + rollingRowNum, person.getIdealPay() / 100.0);
				sheetInCalculation.setCellContent("D" + rollingRowNum, person.getCalcIdealPay() / 100.0);
				sheetInCalculation.setCellContent("E" + rollingRowNum, person.getMaxPay() / 100.0);
				sheetInCalculation.setCellContent("F" + rollingRowNum, person.getAgreedPay() / 100.0);
				sheetInCalculation.setCellContent("G" + rollingRowNum, person.getHadExpense() / 100.0);
				sheetInCalculation.setCellContent("H" + rollingRowNum, person.getTransCosts() / 100.0);
				sheetInCalculation.setCellContent("I" + rollingRowNum, person.getActualTransaction() / 100.0);
				sheetInCalculation.setCellContent("J" + rollingRowNum, person.getNights());
				if (person.getNights() == 0) {
					sheetInCalculation.setCellContent("K" + rollingRowNum, "");
				} else {
					sheetInCalculation.setCellContent("K" + rollingRowNum, person.getAgreedPay() / (person.getNights() * 100.0));
				}
				rollingRowNum++;

				// set actual transaction on IN Payments
				Map<String, Record> payInCol = sheetInPayments.getColContents("A");
				for (Map.Entry<String, Record> entry : payInCol.entrySet()) {
					String key = entry.getKey();
					String rowNum = XlsxSheet.nameToRow(key);
					if (StrUtils.strToInt(rowNum) > HEAD_LINE_AMOUNT_IN_PAYMENTS) {
						Record val = entry.getValue();
						if (val == null) {
							continue;
						}
						String name = val.asString();
						name = Person.sanitizeName(name);
						if (person.getName().equals(name)) {
							sheetInPayments.setCellContent("E" + rowNum, person.getActualTransaction() / 100.0);
							sheetInPayments.setCellContent("H" + rowNum, person.getContactMethod());
						}
					}
				}

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
			if (costCounterTransport != 0) {
				statsContent.append("Includes everyone's transport costs: ");
				statsContent.append(FinanceUtils.formatMoney(costCounterTransport) + " €");
				statsContent.append("\r\n");
			}
			statsContent.append("Includes other costs, such as energy, cleaning, etc.: ");
			statsContent.append(FinanceUtils.formatMoney(costCounterOther) + " €");
			statsContent.append("\r\n");
			if (costCounterCredit != 0) {
				if (costCounterCredit > 0) {
					statsContent.append("Includes credit set aside for next time: ");
					statsContent.append(FinanceUtils.formatMoney(costCounterCredit) + " €");
				} else {
					statsContent.append("Reduced by credit that was set aside last time: ");
					statsContent.append(FinanceUtils.formatMoney(-costCounterCredit) + " €");
				}
				statsContent.append("\r\n");
			}
		}
		statsContent.append("\r\n");
		statsContent.append("Amount of people: " + amountOfPeople + " (" + amountOfPeopleZeroNights + " " +
			(amountOfPeopleZeroNights == 1 ? "person" : "people") +
			" cancelled after the deadline and " +
			(amountOfPeopleZeroNights == 1 ? "is" : "are") +
			" included in the stats)");
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
			(sortedPayList.get(peopleAmountDiv10 - 1) + sortedPayList.get(peopleAmountDiv10)) / 2
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
		if (amountOfNights > 0) {
			statsContent.append("Average ideal payment per night per person: ");
			statsContent.append(FinanceUtils.formatMoney(idealPayCounter / amountOfNights) + " €");
			statsContent.append("\r\n");
			statsContent.append("Average maximum payment per night per person: ");
			statsContent.append(FinanceUtils.formatMoney(maxPayCounter / amountOfNights) + " €");
			statsContent.append("\r\n");
			statsContent.append("Average total payment per night per person: ");
			statsContent.append(FinanceUtils.formatMoney(costCounterSum / amountOfNights) + " €");
			statsContent.append("\r\n");
			statsContent.append("Average food-only payment per night per person: ");
			statsContent.append(FinanceUtils.formatMoney(costCounterFood / amountOfNights) + " €");
			statsContent.append("\r\n");
		}
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

		boolean found = true;

		while (found) {
			found = false;

			int start = outContent.indexOf(fromLine + "\r\n");
			int end = outContent.indexOf(toLine + "\r\n");
			if ((start > -1) && (end > start)) {
				outContent = outContent.substring(0, start) + outContent.substring(end + toLine.length() + 2);
				found = true;
			}

			start = outContent.indexOf(fromLine + "\n");
			end = outContent.indexOf(toLine + "\n");
			if ((start > -1) && (end > start)) {
				outContent = outContent.substring(0, start) + outContent.substring(end + toLine.length() + 1);
				found = true;
			}
		}

		return outContent;
	}

	private static String removeFromTo(String outContent, String from, String to) {

		while (true) {
			int start = outContent.indexOf(from);
			int end = outContent.indexOf(to);
			if ((start > -1) && (end > start)) {
				outContent = outContent.substring(0, start) + outContent.substring(end + to.length());
			} else {
				return outContent;
			}
		}
	}

}
