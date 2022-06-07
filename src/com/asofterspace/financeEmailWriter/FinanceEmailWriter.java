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

	public final static String PROGRAM_TITLE = "FinanceEmailWriter";
	public final static String VERSION_NUMBER = "0.0.0.1(" + Utils.TOOLBOX_VERSION_NUMBER + ")";
	public final static String VERSION_DATE = "6. July 2022 - 6. July 2022";


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
			Integer idealPay = FinanceUtils.parseMoney(content.get(2));
			Integer maxPay = FinanceUtils.parseMoney(content.get(3));
			Integer agreedPay = FinanceUtils.parseMoney(content.get(5));
			Integer hadExpense = FinanceUtils.parseMoney(content.get(6));
			Integer actualTransaction = FinanceUtils.parseMoney(content.get(7));
			Integer nights = StrUtils.strToInt(content.get(8));
			String arrivalDate = content.get(11);
			String leaveDate = content.get(12);

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
			outContent = StrUtils.replaceAll(outContent, "(CONTACT_METHOD)", contactMethod);
			outContent = StrUtils.replaceAll(outContent, "(NAME)", name);
			outContent = StrUtils.replaceAll(outContent, "(IDEAL_PAY)", FinanceUtils.formatMoney(idealPay) + " € (" + FinanceUtils.formatMoney(idealPay / nights) + " € per night)");
			outContent = StrUtils.replaceAll(outContent, "(MAX_PAY)", FinanceUtils.formatMoney(maxPay) + " € (" + FinanceUtils.formatMoney(maxPay / nights) + " € per night)");
			outContent = StrUtils.replaceAll(outContent, "(AGREED_PAY)", FinanceUtils.formatMoney(agreedPay) + " €");
			outContent = StrUtils.replaceAll(outContent, "(HAD_EXPENSE)", FinanceUtils.formatMoney(hadExpense) + " €");
			outContent = StrUtils.replaceAll(outContent, "(ACTUAL_TRANSACTION)", FinanceUtils.formatMoney(actualTransaction) + " €");
			outContent = StrUtils.replaceAll(outContent, "(NIGHTS)", ""+nights);
			outContent = StrUtils.replaceAll(outContent, "(ARRIVAL_DATE)", arrivalDate);
			outContent = StrUtils.replaceAll(outContent, "(LEAVE_DATE)", leaveDate);

			TextFile outFile = new TextFile(outputDir, uniqueName + ".txt");
			outFile.saveContent(outContent);

			content = inputFile.getContentLineInColumns();
		}

	}
}
