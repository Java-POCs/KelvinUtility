package com.map.excel.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.map.excel.model.HeatMapDashBoard;
import com.map.excel.model.RowData;
import com.map.excel.service.XlsService;
import com.map.excel.util.ExcelActivity;

@RestController
public class GetDataFromExcelController {

	@Autowired
	private ExcelActivity excelActivity;

	@Autowired
	private XlsService objXlsService;

	@CrossOrigin(origins = "*", allowedHeaders = "*")
	@PostMapping(value = "/api/upload", produces = "application/json")
	public ResponseEntity<?> uploadFile(@RequestParam("file1") MultipartFile file1,
			@RequestParam("file2") MultipartFile file2, @RequestParam("file3") MultipartFile file3,
			@RequestParam("file4") MultipartFile file4) {
		if (file1.isEmpty() || file2.isEmpty() || file3.isEmpty()) {
			return new ResponseEntity<String>("please select a file!", HttpStatus.NOT_FOUND);
		}
		Map<Integer, String> txCodeSheetHeaderMap = new HashMap<Integer, String>();
		txCodeSheetHeaderMap.put(0, "Transaction code");
		txCodeSheetHeaderMap.put(1, "Sum of Tx Count");
		txCodeSheetHeaderMap.put(2, "Sum of Steps Count");
		txCodeSheetHeaderMap.put(3, "Sum of GUITIME");
		txCodeSheetHeaderMap.put(4, "Roles");
		txCodeSheetHeaderMap.put(5, "Users");
		txCodeSheetHeaderMap.put(6, "Transaction Description");
		txCodeSheetHeaderMap.put(7, "Line of Business");
		txCodeSheetHeaderMap.put(8, "80/20 for Sum of Tx Count");
		txCodeSheetHeaderMap.put(9, "80/20 for Sum of of Steps Count");
		txCodeSheetHeaderMap.put(10, "80/20 for Sum of GUI Time");
		txCodeSheetHeaderMap.put(11, "80/20 for Roles");
		txCodeSheetHeaderMap.put(12, "80/20 for Users");
		String destinationPath = "/tempFolderForFiles";
		List<HeatMapDashBoard> lmdb = null;
		try {
			String extensionOfFile1 = excelActivity.getExtension(file1.getOriginalFilename());
			String extensionOfFile2 = excelActivity.getExtension(file2.getOriginalFilename());
			String extensionOfFile3 = excelActivity.getExtension(file3.getOriginalFilename());
			String extensionOfFile4 = excelActivity.getExtension(file4.getOriginalFilename());
			if ((extensionOfFile1.equalsIgnoreCase("xls") || extensionOfFile1.equalsIgnoreCase("xlsx"))
					&& (extensionOfFile2.equalsIgnoreCase("xls") || extensionOfFile2.equalsIgnoreCase("xlsx"))
					&& (extensionOfFile3.equalsIgnoreCase("xls") || extensionOfFile3.equalsIgnoreCase("xlsx"))
					&& (extensionOfFile4.equalsIgnoreCase("xls") || extensionOfFile4.equalsIgnoreCase("xlsx"))) {
				String txFilePath = excelActivity.saveUploadedFiles(file1);
				String aggrUserFilePath = excelActivity.saveUploadedFiles(file2);
				String agr_1251FilePath = excelActivity.saveUploadedFiles(file3);
				String descriptionPath = excelActivity.saveUploadedFiles(file4);
				List<String> errors = excelFileValidation(txFilePath, aggrUserFilePath, agr_1251FilePath);
				if (errors.size() > 0) {
					this.excelActivity.deleteFile(txFilePath);
					this.excelActivity.deleteFile(aggrUserFilePath);
					this.excelActivity.deleteFile(agr_1251FilePath);
					this.excelActivity.deleteFile(descriptionPath);
					return new ResponseEntity<List<String>>(errors, HttpStatus.BAD_REQUEST);
				}

				Map<String, RowData> txCodeDataMap = objXlsService.readAndConsolidateTxCodeData(txFilePath);
				Map<String, Set<String>> aggrUserDataMap = objXlsService
						.readAndConsolidateAggrUserData(aggrUserFilePath);
				Map<String, RowData> agr_1251DataMap = objXlsService.readAndConsolidateAggr1251UserData(
						agr_1251FilePath, aggrUserDataMap, txCodeDataMap, descriptionPath);

				HSSFWorkbook workbook;
				String excelFileName = System.currentTimeMillis() + "";// name of excel file
				File file = new File(destinationPath + "/" + excelFileName + ".xls");
				if (file.exists() == false) {
					workbook = new HSSFWorkbook();
				} else {
					try (InputStream is = new FileInputStream(file)) {
						workbook = new HSSFWorkbook(is);
					}
				}

				lmdb = objXlsService.prepareExcel(destinationPath, txCodeSheetHeaderMap, agr_1251DataMap, workbook,
						descriptionPath);
				FileOutputStream fileOut = new FileOutputStream(file);
				workbook.write(fileOut);
				fileOut.flush();
				fileOut.close();
				this.excelActivity.deleteFile(txFilePath);
				this.excelActivity.deleteFile(aggrUserFilePath);
				this.excelActivity.deleteFile(agr_1251FilePath);
				this.excelActivity.deleteFile(descriptionPath);

			} else {
				return new ResponseEntity<String>("The uploaded file is not an Excel file", HttpStatus.BAD_REQUEST);
			}
		} catch (Exception e) {
			e.printStackTrace();
			return new ResponseEntity<>(HttpStatus.BAD_REQUEST);
		}

		return new ResponseEntity<List<HeatMapDashBoard>>(lmdb, HttpStatus.OK);
	}

	public List<String> excelFileValidation(String file1, String file2, String file3) {
		List<String> errorMessages = new ArrayList<>();
		List<String> headerForFileOne = new ArrayList<>();
		headerForFileOne.add("ENTRY_ID");
		headerForFileOne.add("COUNT");
		headerForFileOne.add("LUW_COUNT");
		headerForFileOne.add("GUITIME");
		List<String> headerForFileTwo = new ArrayList<>();
		headerForFileTwo.add("AGR_NAME");
		headerForFileTwo.add("UNAME");
		List<String> headerForFileThree = new ArrayList<>();
		headerForFileThree.add("AGR_NAME");
		headerForFileThree.add("LOW");
		try {
			errorMessages.addAll(validatingHeader(file1, headerForFileOne));
			errorMessages.addAll(validatingHeader(file2, headerForFileTwo));
			errorMessages.addAll(validatingHeader(file3, headerForFileThree));
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		}

		return errorMessages;
	}

	public List<String> validatingHeader(String fileName, List<String> header)
			throws EncryptedDocumentException, IOException {
		List<String> errorMessage = new ArrayList<>();
		File file=new File(fileName);
		Workbook wb = WorkbookFactory.create(file);
		Sheet sheet = wb.getSheetAt(0);
		Row headerRow = sheet.getRow(0);
		if (headerRow == null)
			errorMessage.add("No Header found in file " + file.getName());
		else {
			Iterator<Cell> cells = headerRow.cellIterator();
			List<String> cellValue = new ArrayList<String>();
			while (cells.hasNext()) {
				Cell cell = (Cell) cells.next();
				if (cell.getCellTypeEnum() == CellType.STRING)
					cellValue.add(cell.getStringCellValue());
			}
			if (cellValue.size() > 0) {
				List<String> foundHeader = new ArrayList<>();
				for (String head : header) {
					if (cellValue.contains(head))
						foundHeader.add(head);
				}
				header.removeAll(foundHeader);
			} else
				errorMessage.add("No Header found in file " + file.getName());
			if (header.size() > 0) {
				for (String head : header) {
					errorMessage.add("Column " + head + " not found in file " + file.getName());
				}
			}
		}
		wb.close();
		return errorMessage;

	}

}
