package com.map.excel.service;

import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import com.map.excel.model.HeatMapDashBoard;
import com.map.excel.model.RowData;
import com.map.excel.util.XlsReaderUtils;
import com.map.excel.util.XlsWriterUtils;

@Component
public class XlsService {

	@Autowired
	private XlsReaderUtils objXlsReader;

	@Autowired
	private XlsWriterUtils objXlsWriter;

	public Map<String, RowData> readAndConsolidateTxCodeData(String inputXlsPath) throws Exception {
		String groupByColumn = "ENTRY_ID";
		Map<String, Integer> headerColIdxMap = new HashMap<String, Integer>();
		Map<String, RowData> dataMap = new LinkedHashMap<String, RowData>();
		// Workbook wb = objXlsReader.getXlsWorkbook(inputXlsPath);
		Workbook wb = objXlsReader.getXlsWorkbook(inputXlsPath);
		int sheetNo = 0; // Sheet To read
		int headerRowNum = 0;
		// Set the indexes of required columns
		mapColIndex(headerColIdxMap, wb, sheetNo, headerRowNum, "ENTRY_ID");
		mapColIndex(headerColIdxMap, wb, sheetNo, headerRowNum, "COUNT");
		mapColIndex(headerColIdxMap, wb, sheetNo, headerRowNum, "LUW_COUNT");
		mapColIndex(headerColIdxMap, wb, sheetNo, headerRowNum, "GUITIME");

		Map<Integer, String> colNameIdxMap = createIdxColMap(headerColIdxMap);

		dataMap = objXlsReader.getColVals(wb, sheetNo, headerColIdxMap, colNameIdxMap, groupByColumn);
		return dataMap;

	}

	public Map<String, RowData> readAndConsolidateAggr1251UserData(String inputFilePath,
			Map<String, Set<String>> aggrUserDataMap, Map<String, RowData> txCodeDataMap, String destinationPath)
			throws Exception {
		Workbook wb = objXlsReader.getXlsWorkbook(inputFilePath);
		int sheetNo = 0; // Sheet To read
		Map<String, RowData> hmRowData = objXlsReader.readAgr1251Sheet(wb, sheetNo, aggrUserDataMap, txCodeDataMap);
		return hmRowData;
	}

	public Map<String, Set<String>> readAndConsolidateAggrUserData(String aggrUserFilePath) throws Exception {
		Workbook wb = objXlsReader.getXlsWorkbook(aggrUserFilePath);
		int sheetNo = 0; // Sheet To read
		Map<String, Set<String>> aggrUserDataMap = objXlsReader.readAgrSheet(wb, sheetNo);
		return aggrUserDataMap;
	}

	private Map<Integer, String> createIdxColMap(Map<String, Integer> headerColIdxMap) {
		Map<Integer, String> colNameIdxMap = new HashMap<>();
		Iterator<String> itr = headerColIdxMap.keySet().iterator();
		while (itr.hasNext()) {
			String colName = itr.next();
			int colIdx = headerColIdxMap.get(colName);
			colNameIdxMap.put(colIdx, colName);
		}
		return colNameIdxMap;
	}

	private List<String> getDistinctValsForCol(Workbook wb, int sheetNo, Integer colIndex) {
		List<String> vals = objXlsReader.getAllValsForCol(wb, sheetNo, colIndex);
		List<String> distinctVals = vals.stream().distinct().collect(Collectors.toList());
		return distinctVals;
	}

	private void mapColIndex(Map<String, Integer> headerColIdxMap, Workbook wb, int sheetNo, int headerRowNum,
			String matchColName) throws Exception {
		int matchedColNum = objXlsReader.getColNoForColName(wb, sheetNo, headerRowNum, matchColName);
		if (matchedColNum < 0)
			throw new Exception("Column Name [" + matchColName + "] not found in sheet");
		headerColIdxMap.put(matchColName, matchedColNum);
	}

	public List<HeatMapDashBoard> prepareExcel(String destinationPath, Map<Integer, String> headerMap,
			Map<String, RowData> dataMap, Workbook workbook, String descriptionPath) throws IOException {
		String sheetName = "Sheet1";// name of sheet
		objXlsWriter.writeHeadersToXls(workbook, sheetName, headerMap, 0);
		List<HeatMapDashBoard> lmdb = objXlsWriter.writeDataToXls(workbook, sheetName, dataMap, descriptionPath);
		return lmdb;
	}

}
