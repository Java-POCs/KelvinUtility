package com.map.excel.util;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import com.map.excel.model.HeatMapDashBoard;
import com.map.excel.model.RowData;
import com.map.excel.model.RowDataForCumalative;
import com.map.excel.model.SortByGuiTime;
import com.map.excel.model.SortByRole;
import com.map.excel.model.SortByStepsCount;
import com.map.excel.model.SortByTxCount;
import com.map.excel.model.SortByUser;
import com.monitorjbl.xlsx.StreamingReader;

@Component
public class XlsWriterUtils {

	private DecimalFormat df2 = new DecimalFormat("#.##");

	//private Map<String, List<String>> getDescriptionOfTrxCode = null;

	public Workbook writeHeadersToXls(Workbook wb, String sheetName, Map<Integer, String> dataMap, int rowNum)
			throws IOException {
		Sheet sheet = null;
		// Check if the workbook is empty or not
		if (wb.getNumberOfSheets() != 0) {
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				if (wb.getSheetName(i).equals(sheetName)) {
					sheet = wb.getSheet(sheetName);
				} else
					sheet = wb.createSheet(sheetName);
			}
		} else {
			// Create new sheet to the workbook if empty
			sheet = wb.createSheet(sheetName);
		}

		Row row = sheet.createRow(rowNum);
		Iterator<Integer> itrRowData = dataMap.keySet().iterator();
		while (itrRowData.hasNext()) {
			int key = itrRowData.next();
			String val = dataMap.get(key);
			Cell cell = row.createCell(key);
			cell.setCellValue(val);
		}
		return wb;
	}

	public List<HeatMapDashBoard> writeDataToXls(Workbook wb, String sheetName, Map<String, RowData> dataMap,
			String descriptionFilePath) throws IOException {
		Sheet sheet = wb.getSheet(sheetName);
		//if (getDescriptionOfTrxCode == null) {
		Map<String, List<String>>	getDescriptionOfTrxCode = readDataForDescriptionFromExcel(descriptionFilePath);
		//}
		int rowNum = 1;
		BigDecimal totalSumOfTxCount = new BigDecimal(0);
		BigDecimal totalSumOfStepCount = new BigDecimal(0);
		BigDecimal totalSumOfGuiTime = new BigDecimal(0);
		BigDecimal totalSumOfRoles = new BigDecimal(0);
		BigDecimal totalSumOfUsers = new BigDecimal(0);
		Iterator<String> itrDataMap = dataMap.keySet().iterator();
		while (itrDataMap.hasNext()) {
			String description = null, lob = null;
			String key = itrDataMap.next();
			RowData objRow = dataMap.get(key);
			Row row = sheet.createRow(rowNum);

			Cell txCodeCell = row.createCell(0);
			txCodeCell.setCellType(CellType.STRING);
			txCodeCell.setCellValue(objRow.getTxCode());

			Cell txCountCell = row.createCell(1);
			txCountCell.setCellType(CellType.NUMERIC);
			txCountCell.setCellValue(objRow.getBgSumTxCount().doubleValue());

			Cell txStepCountCell = row.createCell(2);
			txStepCountCell.setCellType(CellType.NUMERIC);
			txStepCountCell.setCellValue(objRow.getBgSumStepCount().doubleValue());

			Cell txGuiTimeCell = row.createCell(3);
			txGuiTimeCell.setCellType(CellType.NUMERIC);
			txGuiTimeCell.setCellValue(objRow.getBgSumGuiTime().doubleValue());

			Cell txRoles = row.createCell(4);
			txRoles.setCellType(CellType.NUMERIC);
			txRoles.setCellValue(objRow.getRoles().size());

			Cell txUsers = row.createCell(5);
			txUsers.setCellType(CellType.NUMERIC);
			txUsers.setCellValue(objRow.getUsers().size());

			if (getDescriptionOfTrxCode.get(key) == null) {
				description = "Others";
				lob = "O";
			} else {
				List<String> mapList = getDescriptionOfTrxCode.get(key);
				if (mapList.size() > 0) {
					if (mapList.size() > 1) {
						lob = mapList.get(1);
					} else
						lob = "O";
					description = mapList.get(0);
				} else {
					description = "Others";
					lob = "O";
				}
			}
			Cell descriptionOfTxCode = row.createCell(6);
			descriptionOfTxCode.setCellType(CellType.STRING);
			descriptionOfTxCode.setCellValue(description);

			Cell lineOfBusiness = row.createCell(7);
			lineOfBusiness.setCellType(CellType.STRING);
			lineOfBusiness.setCellValue(lob);

			rowNum++;
			totalSumOfTxCount = totalSumOfTxCount.add(objRow.getBgSumTxCount());
			totalSumOfStepCount = totalSumOfStepCount.add(objRow.getBgSumStepCount());
			totalSumOfGuiTime = totalSumOfGuiTime.add(objRow.getBgSumGuiTime());
			totalSumOfRoles = totalSumOfRoles.add(BigDecimal.valueOf(objRow.getRoles().size()));
			totalSumOfUsers = totalSumOfUsers.add(BigDecimal.valueOf(objRow.getUsers().size()));

		}

		List<HeatMapDashBoard> lmdb=addPercentage(wb, totalSumOfTxCount, totalSumOfStepCount, totalSumOfGuiTime, totalSumOfRoles, totalSumOfUsers,
				dataMap);
		return lmdb;
	}

	public List<HeatMapDashBoard> addPercentage(Workbook wb, BigDecimal totalSumOfTxCount, BigDecimal totalSumOfStepCount,
			BigDecimal totalSumOfGuiTime, BigDecimal totalSumOfRoles, BigDecimal totalSumOfUsers,
			Map<String, RowData> dataMap) {

		Set<RowDataForCumalative> forTrxCount = new TreeSet<>(new SortByTxCount());
		Set<RowDataForCumalative> forGuiTime = new TreeSet<>(new SortByGuiTime());
		Set<RowDataForCumalative> forStepsCount = new TreeSet<>(new SortByStepsCount());
		Set<RowDataForCumalative> forRoles = new TreeSet<>(new SortByRole());
		Set<RowDataForCumalative> forUser = new TreeSet<>(new SortByUser());
		Iterator<String> itrDataMap = dataMap.keySet().iterator();
		while (itrDataMap.hasNext()) {
			String key = itrDataMap.next();
			RowData objRow = dataMap.get(key);

			RowDataForCumalative rdfc = new RowDataForCumalative();
			rdfc.setTrxCode(key);
			rdfc.setIndividualPercentageForGuiTime(getPercantage(objRow.getBgSumGuiTime(), totalSumOfGuiTime));
			rdfc.setIndividualPercentageForStepCount(getPercantage(objRow.getBgSumStepCount(), totalSumOfStepCount));
			rdfc.setIndividualPercentageForTxCount(getPercantage(objRow.getBgSumTxCount(), totalSumOfTxCount));
			rdfc.setIndividualPercentageForRoles(
					getPercantage(BigDecimal.valueOf(objRow.getRoles().size()), totalSumOfRoles));
			rdfc.setIndividualPercentageForUsers(
					getPercantage(BigDecimal.valueOf(objRow.getUsers().size()), totalSumOfUsers));

			forGuiTime.add(rdfc);
			forTrxCount.add(rdfc);
			forStepsCount.add(rdfc);
			forRoles.add(rdfc);
			forUser.add(rdfc);
		}

		Map<String, RowDataForCumalative> mapGuiTime = settingCumulativeGUITime(forGuiTime);
		Map<String, RowDataForCumalative> mapTxCount = settingCumulativeTrxCount(forTrxCount);
		Map<String, RowDataForCumalative> mapStepsCount = settingCumulativeStepsCount(forStepsCount);
		Map<String, RowDataForCumalative> mapRoles = settingCumulativeRoles(forRoles);
		Map<String, RowDataForCumalative> mapUsers = settingCumulativeUser(forUser);

		List<HeatMapDashBoard> lmdb=writeDataWithPercentage(wb, mapGuiTime, mapTxCount, mapStepsCount, mapRoles, mapUsers);
		return lmdb;

	}

	public List<HeatMapDashBoard> writeDataWithPercentage(Workbook wb, Map<String, RowDataForCumalative> mapGuiTime,
			Map<String, RowDataForCumalative> mapTxCount, Map<String, RowDataForCumalative> mapStepsCount,
			Map<String, RowDataForCumalative> mapRoles, Map<String, RowDataForCumalative> mapUsers) {
		Sheet sheet = wb.getSheetAt(0);
		List<HeatMapDashBoard> lmdb=new ArrayList<>();
		for (int rowIdx = 1; rowIdx < sheet.getPhysicalNumberOfRows(); rowIdx++) {
			HeatMapDashBoard hmdb=new HeatMapDashBoard();
			Row row = sheet.getRow(rowIdx);
			String key = row.getCell(0).getStringCellValue();
			hmdb.setTrxCode(key);
			hmdb.setTrxCount(row.getCell(1).getNumericCellValue());
			hmdb.setStepsCount(row.getCell(2).getNumericCellValue());
			hmdb.setGuiTime(row.getCell(3).getNumericCellValue());
			hmdb.setRoles(row.getCell(4).getNumericCellValue());
			hmdb.setUser(row.getCell(5).getNumericCellValue());
			hmdb.setDescription(row.getCell(6).getStringCellValue());
			hmdb.setLineOfBusiness(row.getCell(7).getStringCellValue());

			Cell cumulativeTxCountCell = row.createCell(8);
			cumulativeTxCountCell.setCellType(CellType.NUMERIC);
			double cumTrxCount=mapTxCount.get(key).getCumlativeValueForTxCount();
			cumulativeTxCountCell.setCellValue(cumTrxCount);
			hmdb.setCumlativeValueForTrxCount(cumTrxCount);

			Cell cumulativeTxStepCountCell = row.createCell(9);
			cumulativeTxStepCountCell.setCellType(CellType.NUMERIC);
			double cumStepsCount=mapStepsCount.get(key).getCumlativeValueForStepCount();
			cumulativeTxStepCountCell.setCellValue(cumStepsCount);
			hmdb.setCumlativeValueForStepsCount(cumStepsCount);

			Cell cumulativeTxGuiTimeCell = row.createCell(10);
			cumulativeTxGuiTimeCell.setCellType(CellType.NUMERIC);
			double cumGuiTime=mapGuiTime.get(key).getCumlativeValueForGuiTime();
			cumulativeTxGuiTimeCell.setCellValue(cumGuiTime);
			hmdb.setCumlativeValueForGuiTime(cumGuiTime);

			Cell cumulativeRolesCell = row.createCell(11);
			cumulativeRolesCell.setCellType(CellType.NUMERIC);
			double cumRoles=mapRoles.get(key).getCumlativeValueForRoles();
			cumulativeRolesCell.setCellValue(cumRoles);
			hmdb.setCumlativeValueForRoles(cumRoles);

			Cell cumulativeUser = row.createCell(12);
			cumulativeUser.setCellType(CellType.NUMERIC);
			double cumUsers=mapUsers.get(key).getCumlativeValueForUsers();
			cumulativeUser.setCellValue(cumUsers);
			hmdb.setCumlativeValueForUser(cumUsers);
			lmdb.add(hmdb);
		}

		return lmdb;
	}

	public double getPercantage(BigDecimal count, BigDecimal totalSum) {
		double cnt = count.doubleValue();
		double totSum = totalSum.doubleValue();
		if (totSum == 0)
			return 0;

		return Double.parseDouble(this.df2.format((cnt / totSum) * 100));
	}

	public Map<String, RowDataForCumalative> settingCumulativeGUITime(Set<RowDataForCumalative> hmdb2) {
		Map<String, RowDataForCumalative> tMap = new HashMap<>();
		double cumulativeValueForGUITime = 0;
		for (RowDataForCumalative hmd : hmdb2) {
			cumulativeValueForGUITime += hmd.getIndividualPercentageForGuiTime();
			hmd.setCumlativeValueForGuiTime(Double.parseDouble(this.df2.format(cumulativeValueForGUITime)));
			tMap.put(hmd.getTrxCode(), hmd);
		}
		return tMap;
	}

	public Map<String, RowDataForCumalative> settingCumulativeTrxCount(Set<RowDataForCumalative> hmdb2) {
		Map<String, RowDataForCumalative> tMap = new HashMap<>();
		double cumulativeValueForTrxCount = 0;
		for (RowDataForCumalative hmd : hmdb2) {
			cumulativeValueForTrxCount += hmd.getIndividualPercentageForTxCount();
			hmd.setCumlativeValueForTxCount(Double.parseDouble(this.df2.format(cumulativeValueForTrxCount)));
			tMap.put(hmd.getTrxCode(), hmd);
		}
		return tMap;
	}

	public Map<String, RowDataForCumalative> settingCumulativeStepsCount(Set<RowDataForCumalative> hmdb2) {
		Map<String, RowDataForCumalative> tMap = new HashMap<>();
		double cumulativeValueForStepsCount = 0;
		for (RowDataForCumalative hmd : hmdb2) {
			cumulativeValueForStepsCount += hmd.getIndividualPercentageForStepCount();
			hmd.setCumlativeValueForStepCount(Double.parseDouble(this.df2.format(cumulativeValueForStepsCount)));
			tMap.put(hmd.getTrxCode(), hmd);
		}
		return tMap;
	}

	public Map<String, RowDataForCumalative> settingCumulativeRoles(Set<RowDataForCumalative> hmdb2) {
		Map<String, RowDataForCumalative> tMap = new HashMap<>();
		double cumulativeValueForRoles = 0;
		for (RowDataForCumalative hmd : hmdb2) {
			cumulativeValueForRoles += hmd.getIndividualPercentageForRoles();
			hmd.setCumlativeValueForRoles(Double.parseDouble(this.df2.format(cumulativeValueForRoles)));
			tMap.put(hmd.getTrxCode(), hmd);
		}
		return tMap;
	}

	public Map<String, RowDataForCumalative> settingCumulativeUser(Set<RowDataForCumalative> hmdb2) {
		Map<String, RowDataForCumalative> tMap = new HashMap<>();
		double cumulativeValueForUsers = 0;
		for (RowDataForCumalative hmd : hmdb2) {
			cumulativeValueForUsers += hmd.getIndividualPercentageForUsers();
			if (cumulativeValueForUsers > 100)
				cumulativeValueForUsers = 100;
			hmd.setCumlativeValueForUsers(Double.parseDouble(this.df2.format(cumulativeValueForUsers)));
			tMap.put(hmd.getTrxCode(), hmd);
		}
		return tMap;
	}

	public Map<String, List<String>> readDataForDescriptionFromExcel(String fileNameWithPath) {
		Map<String, List<String>> getDescription = new HashMap<String, List<String>>();
		String cellValue = null;
		String key = null;
		Path excelFileToRead = null;
		InputStream is = null;
		Workbook workbook = null;
		Sheet firstSheet = null;
		try {
			excelFileToRead = Paths.get(fileNameWithPath);
			is = Files.newInputStream(excelFileToRead);
			workbook = StreamingReader.builder().rowCacheSize(100).bufferSize(4096).open(is);
			firstSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = firstSheet.iterator();
			int noOfRows = 0;
			while (iterator.hasNext()) {
				List<String> lis = new ArrayList<String>();
				Row _nextRow = iterator.next();
				Iterator<Cell> cellIterator = _nextRow.cellIterator();
				if (noOfRows > 0) {
					int i = 0;
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						switch (cell.getCellTypeEnum()) {
						case STRING:
							cellValue = cell.getStringCellValue();
							if (i == 0)
								key = cellValue;
							else
								lis.add(cellValue);
							break;
						default:
							break;
						}
						i++;
					}
					getDescription.put(key, lis);
				}
				noOfRows++;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return getDescription;
	}

}
