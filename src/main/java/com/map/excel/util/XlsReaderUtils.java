package com.map.excel.util;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Component;

import com.map.excel.model.RowData;

@Component
public class XlsReaderUtils {

	public Workbook getXlsWorkbook(String inputFile) throws IOException {
		// InputStream ExcelFileToRead = new FileInputStream(inputFile);
		// Workbook wb = new Workbook(ExcelFileToRead);
		// Workbook workbook =
		// StreamingReader.builder().rowCacheSize(100).bufferSize(4096).open(ExcelFileToRead);
		Workbook wb = WorkbookFactory.create(new File(inputFile));
		return wb;
	}

	public int getColNoForColName(Workbook wb, int sheetNo, int headerRowNum, String colName) {
		int matchedColNum = -1;
		Sheet sheet = wb.getSheetAt(sheetNo);
		Row headerRow = sheet.getRow(headerRowNum);
		Iterator<Cell> cells = headerRow.cellIterator();
		while (cells.hasNext()) {
			Cell cell = (Cell) cells.next();
			if (colName.equalsIgnoreCase(cell.getStringCellValue())) {
				matchedColNum = cell.getColumnIndex();
				break;
			}
		}
		return matchedColNum;
	}

	public List<String> getAllValsForCol(Workbook wb, int sheetNo, int colIndex) {
		List<String> alVals = new ArrayList<String>();
		Sheet sheet = wb.getSheetAt(sheetNo);
		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			Cell cell = row.getCell(colIndex);
			alVals.add(cell.getStringCellValue());
		}
		return alVals;
	}

	public Map<String, Set<String>> readAgrSheet(Workbook wb, int sheetNo) throws Exception {
		Map<String, Set<String>> hmRowData = new HashMap<String, Set<String>>();
		Map<String, Integer> hmColIdxMap = new LinkedHashMap<String, Integer>();
		Set<String> users = new HashSet<>();
		Sheet sheet = wb.getSheetAt(sheetNo);
		Row headerRow = sheet.getRow(0);
		Iterator<Cell> cells = headerRow.cellIterator();
		while (cells.hasNext()) {
			Cell headerCell = (Cell) cells.next();
			hmColIdxMap.put(headerCell.getStringCellValue(), headerCell.getColumnIndex());
		}

		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			String role = getStringVal(hmColIdxMap, row, "AGR_NAME");
			String user = getStringVal(hmColIdxMap, row, "UNAME");
			if (role != null && user != null) {
				role = role.trim();
				if (hmRowData.containsKey(role))
					users = hmRowData.get(role);
				else
					users = new HashSet<>();
				// if (!users.contains(user))
				users.add(user.trim());
				hmRowData.put(role, users);
			}
		}

		return hmRowData;
	}

	public Map<String, RowData> readAgr1251Sheet(Workbook wb, int sheetNo, Map<String, Set<String>> aggrUserDataMap,
			Map<String, RowData> txCodeDataMap) throws Exception {
		Map<String, RowData> hmRowData = new HashMap<String, RowData>();
		Map<String, Integer> hmColIdxMap = new LinkedHashMap<String, Integer>();
		RowData objRowData = new RowData();
		List<String> roles = new ArrayList<>();
		Sheet sheet = wb.getSheetAt(sheetNo);
		Row headerRow = sheet.getRow(0);
		Iterator<Cell> cells = headerRow.cellIterator();
		while (cells.hasNext()) {
			Cell headerCell = (Cell) cells.next();
			hmColIdxMap.put(headerCell.getStringCellValue(), headerCell.getColumnIndex());
		}

		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			String role = getStringVal(hmColIdxMap, row, "AGR_NAME");
			String txId = getStringVal(hmColIdxMap, row, "LOW");
			if (txId != null && role != null) {
				txId = txId.trim();
				role = role.trim();
				if (txCodeDataMap.containsKey(txId)) {
					Set<String> user = new HashSet<>();
					if (hmRowData.containsKey(txId)) {
						objRowData = hmRowData.get(txId);
						user = objRowData.getUsers();
					} else
						objRowData = new RowData();
					objRowData.setTxCode(txId);
					roles = objRowData.getRoles();
					roles.add(role);
					objRowData.setRoles(roles);

					if (aggrUserDataMap.containsKey(role)) {
						Set<String> user1 = aggrUserDataMap.get(role);
						user.addAll(user1);
						objRowData.setUsers(user);
					} else {
						objRowData.setUsers(user);
					}
					RowData rowDataTxCodeSheet = txCodeDataMap.get(txId);
					objRowData.setBgSumGuiTime(rowDataTxCodeSheet.getBgSumGuiTime());
					objRowData.setBgSumStepCount(rowDataTxCodeSheet.getBgSumStepCount());
					objRowData.setBgSumTxCount(rowDataTxCodeSheet.getBgSumTxCount());
					/*
					 * if (txCodeDataMap.containsKey(txId)) { RowData rowDataTxCodeSheet =
					 * txCodeDataMap.get(txId);
					 * objRowData.setBgSumGuiTime(rowDataTxCodeSheet.getBgSumGuiTime());
					 * objRowData.setBgSumStepCount(rowDataTxCodeSheet.getBgSumStepCount());
					 * objRowData.setBgSumTxCount(rowDataTxCodeSheet.getBgSumTxCount()); } else {
					 * objRowData.setBgSumGuiTime(new BigDecimal(0));
					 * objRowData.setBgSumStepCount(new BigDecimal(0));
					 * objRowData.setBgSumTxCount(new BigDecimal(0)); }
					 */
					hmRowData.put(txId, objRowData);
					roles = new ArrayList<>();
				}
			}
		}

		return hmRowData;
	}

	private String getStringVal(Map<String, Integer> hmColIdxMap, Row row, String key) throws Exception {
		Cell cell;
		String val = null;
		if (!hmColIdxMap.containsKey(key))
			throw new Exception("Column" + key + " not found in sheet");
		cell = row.getCell(hmColIdxMap.get(key));
		if (cell == null)
			return val;

		switch (cell.getCellTypeEnum()) {
		case NUMERIC:
			val = cell.getNumericCellValue() + "";
			break;
		case STRING:
			val = cell.getStringCellValue();
			break;
		}
		return val;
	}

	public Map<String, RowData> getColVals(Workbook wb, int sheetNo, Map<String, Integer> headerColIdxMap,
			Map<Integer, String> colNameIdxMap, String groupByColumn) {
		Cell cell;
		Map<String, RowData> dataMap = new HashMap<String, RowData>();
		Sheet sheet = wb.getSheetAt(sheetNo);
		int getColNoForTxCount = headerColIdxMap.get("COUNT");
		int getColNoForStepsCount = headerColIdxMap.get("LUW_COUNT");
		int getColNoForGuiTime = headerColIdxMap.get("GUITIME");
		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			Iterator<Cell> cells = row.cellIterator();
			dataLoop: while (cells.hasNext()) {
				cell = (Cell) cells.next();
				int colIdx = cell.getColumnIndex();
				if (colNameIdxMap.containsKey(colIdx)) {
					RowData rowData = null;
					if (colIdx != headerColIdxMap.get(groupByColumn)) {
						Cell txCell = row.getCell(headerColIdxMap.get(groupByColumn));
						String txVal = txCell.getStringCellValue();
						if (!(txVal == null || txVal.trim().equals(""))) {
							txVal = txVal.trim();
							if (txVal.contains(" ") || txVal.contains("	")) {
								if (txVal.charAt(txVal.length() - 1) == 'T'
										|| txVal.charAt(txVal.length() - 1) == 't') {
									txVal = txVal.substring(0, txVal.length() - 1);
									txVal = txVal.replaceAll("\\s+", "");
								} else {
									break dataLoop;
								}
							}

						} else
							break dataLoop;
						if (dataMap.containsKey(txVal))
							rowData = dataMap.get(txVal);
						else {
							rowData = new RowData();
							rowData.setTxCode(txVal);
						}
						BigDecimal bgCellVal = null;
						switch (cell.getCellType()) {
						case NUMERIC:
							bgCellVal = new BigDecimal(cell.getNumericCellValue());
							break;
						case STRING:
							bgCellVal = new BigDecimal(cell.getStringCellValue());
							break;
						default:
							break;

						}

						if (colIdx == getColNoForTxCount) {
							BigDecimal bgTxCount = rowData.getBgSumTxCount();
							bgTxCount = bgTxCount.add(bgCellVal != null ? bgCellVal : new BigDecimal(0));
							rowData.setBgSumTxCount(bgTxCount);
						} else if (colIdx == getColNoForStepsCount) {
							BigDecimal bgSumStepCount = rowData.getBgSumStepCount();
							bgSumStepCount = bgSumStepCount.add(bgCellVal != null ? bgCellVal : new BigDecimal(0));
							rowData.setBgSumStepCount(bgSumStepCount);
						} else if (colIdx == getColNoForGuiTime) {
							BigDecimal bgSumGuiTime = rowData.getBgSumGuiTime();
							bgSumGuiTime = bgSumGuiTime.add(bgCellVal != null ? bgCellVal : new BigDecimal(0));
							rowData.setBgSumGuiTime(bgSumGuiTime);
						}

						/*
						 * switch (colIdx) { case 3: BigDecimal bgTxCount = rowData.getBgSumTxCount();
						 * bgTxCount = bgTxCount.add(bgCellVal != null ? bgCellVal : new BigDecimal(0));
						 * rowData.setBgSumTxCount(bgTxCount); break; case 4: BigDecimal bgSumStepCount
						 * = rowData.getBgSumStepCount(); bgSumStepCount = bgSumStepCount.add(bgCellVal
						 * != null ? bgCellVal : new BigDecimal(0));
						 * rowData.setBgSumStepCount(bgSumStepCount); break; case 10: BigDecimal
						 * bgSumGuiTime = rowData.getBgSumGuiTime(); bgSumGuiTime =
						 * bgSumGuiTime.add(bgCellVal != null ? bgCellVal : new BigDecimal(0));
						 * rowData.setBgSumGuiTime(bgSumGuiTime); break; default: break; }
						 */

						dataMap.put(rowData.getTxCode(), rowData);
					}

				}
			}
		}
		return dataMap;
	}

}
