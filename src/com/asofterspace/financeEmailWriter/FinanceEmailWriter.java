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
import com.asofterspace.toolbox.utils.StrUtils;
import com.asofterspace.toolbox.Utils;

import java.util.HashSet;
import java.util.List;
import java.util.Set;


public class FinanceEmailWriter {

	private static final String NAME = "(NAME)";
	private static final String CONTACT_METHOD = "(CONTACT_METHOD)";
	private static final String IDEAL_PAY = "(IDEAL_PAY)";
	private static final String MAX_PAY = "(MAX_PAY)";
	private static final String AGREED_PAY = "(AGREED_PAY)";
	private static final String END_EXPENSE = "(END_EXPENSE)";
	private static final String BEGIN_EXPENSE = "(BEGIN_EXPENSE)";
	private static final String HAD_EXPENSE = "(HAD_EXPENSE)";
	private static final String ACTUAL_TRANSACTION = "(ACTUAL_TRANSACTION)";
	private static final String NIGHTS = "(NIGHTS)";
	private static final String ARRIVAL_DATE = "(ARRIVAL_DATE)";
	private static final String LEAVE_DATE = "(LEAVE_DATE)";

	public final static String PROGRAM_TITLE = "FinanceEmailWriter";
	public final static String VERSION_NUMBER = "0.0.0.2(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "6. June 2022 - 12. June 2022";


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
		List<String> content = inputFile.getContentLineInColumns();
		while (content != null) {
			String contactMethod = content.get(0).trim();
			String name = content.get(1);
			name = UniversalTextDecoder.decode(name).trim();
			name = StrUtils.removeTrailingPronounsFromName(name);
			Integer idealPay = FinanceUtils.parseMoney(content.get(2));
			Integer maxPay = FinanceUtils.parseMoney(content.get(3));
			Integer agreedPay = FinanceUtils.parseMoney(content.get(5));
			Integer hadExpense = FinanceUtils.parseMoney(content.get(6));
			Integer actualTransaction = FinanceUtils.parseMoney(content.get(7));
			Integer nights = StrUtils.strToInt(content.get(8));
			String arrivalDate = content.get(11);
			String leaveDate = content.get(12);

			if ("".equals(name.trim())) {
				break;
			}

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
				int start = outContent.indexOf(BEGIN_EXPENSE + "\r\n");
				int end = outContent.indexOf(END_EXPENSE + "\r\n");
				if ((start > -1) && (end > -1)) {
					outContent = outContent.substring(0, start) + outContent.substring(end + END_EXPENSE.length() + 2);
				}
				start = outContent.indexOf(BEGIN_EXPENSE + "\n");
				end = outContent.indexOf(END_EXPENSE + "\n");
				if ((start > -1) && (end > -1)) {
					outContent = outContent.substring(0, start) + outContent.substring(end + END_EXPENSE.length() + 1);
				}
			} else {
				outContent = StrUtils.replaceAll(outContent, BEGIN_EXPENSE + "\r\n", "");
				outContent = StrUtils.replaceAll(outContent, END_EXPENSE + "\r\n", "");
				outContent = StrUtils.replaceAll(outContent, BEGIN_EXPENSE + "\n", "");
				outContent = StrUtils.replaceAll(outContent, END_EXPENSE + "\n", "");
			}
			outContent = StrUtils.replaceAll(outContent, HAD_EXPENSE, FinanceUtils.formatMoney(hadExpense) + " €");
			outContent = StrUtils.replaceAll(outContent, ACTUAL_TRANSACTION, FinanceUtils.formatMoney(actualTransaction) + " €");
			outContent = StrUtils.replaceAll(outContent, NIGHTS, ""+nights);
			outContent = StrUtils.replaceAll(outContent, ARRIVAL_DATE, arrivalDate);
			outContent = StrUtils.replaceAll(outContent, LEAVE_DATE, leaveDate);

			TextFile outFile = new TextFile(outputDir, uniqueName + ".txt");
			outFile.saveContent(outContent);

			content = inputFile.getContentLineInColumns();
		}

	}
}
