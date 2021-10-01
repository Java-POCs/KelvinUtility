package com.map.excel.util;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import com.monitorjbl.xlsx.StreamingReader;

@Component
public class ExcelActivity {

	public void deleteFile(String fileNameWithPath) {
		File file = new File(fileNameWithPath);
		if (file.exists())
			file.delete();
	}

	public String getExtension(String fileName) {
		char ch;
		int len;
		if (fileName == null || (len = fileName.length()) == 0 || (ch = fileName.charAt(len - 1)) == '/' || ch == '\\'
				|| // in the case of a directory
				ch == '.') // in the case of . or ..
			return "";
		int dotInd = fileName.lastIndexOf('.'),
				sepInd = Math.max(fileName.lastIndexOf('/'), fileName.lastIndexOf('\\'));
		if (dotInd <= sepInd)
			return "";
		else
			return fileName.substring(dotInd + 1).toLowerCase();
	}

	public Iterator<Row> getRowIterator(String fileNameWithPath) {
		Path excelFileToRead = null;
		InputStream is = null;
		Workbook workbook = null;
		Sheet firstSheet = null;
		Iterator<Row> iterator = null;
		try {
			excelFileToRead = Paths.get(fileNameWithPath);
			is = Files.newInputStream(excelFileToRead);
			workbook = StreamingReader.builder().rowCacheSize(100).bufferSize(4096).open(is);
			firstSheet = workbook.getSheetAt(0);
			iterator = firstSheet.iterator();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return iterator;
	}

	public String saveUploadedFiles(MultipartFile file) throws IOException {
		byte[] bytes = file.getBytes();
		createAFolder("/");
		Path path = Paths.get("/tempFolderForFiles/" + file.getOriginalFilename());
		Files.write(path, bytes);
		return path.toAbsolutePath().toString();
	}

	public void createAFolder(String pathtoCreateFolder) {
		Path path = Paths.get(pathtoCreateFolder + "tempFolderForFiles");
		if (!Files.exists(path)) {
			try {
				Files.createDirectories(path);
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}
