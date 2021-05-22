package com.skazi.excel2json;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.skazi.excel2json.domain.Quiz;

public class Application {

	public static void main(String[] args) {
		// Step 1: Read Excel File into Java List Objects
		List<Quiz> quiz = readExcelFile("C:\\Users\\kazis\\Documents\\Quiz.xlsx");

		// Step 2: Convert Java Objects to JSON String
		String jsonString = convertObjects2JsonString(quiz);

		System.out.println(jsonString);
	}

	/**
	 * Read Excel File into Java List Objects
	 * 
	 * @param filePath
	 * @return
	 */
	private static List<Quiz> readExcelFile(String filePath) {
		try {
			FileInputStream excelFile = new FileInputStream(new File(filePath));
			Workbook workbook = new XSSFWorkbook(excelFile);

			Sheet sheet = workbook.getSheetAt(0);

			List<Quiz> lstQuiz = new ArrayList<Quiz>();
			for (Row currentRow : sheet) {
				if(currentRow.getRowNum()==0) {
    				continue;
    			}
				Quiz quiz = new Quiz();
				quiz.setQue(String.valueOf((Cell) currentRow.getCell(0)));
				if(((Cell) currentRow.getCell(1))!=null)
					quiz.setCode(String.valueOf((Cell) currentRow.getCell(1)));
				else 
					quiz.setCode("");
				String options[] = new String[4];
				options[0] = String.valueOf((Cell) currentRow.getCell(2));
				options[1] = (String.valueOf((Cell) currentRow.getCell(3)));
				options[2] = (String.valueOf((Cell) currentRow.getCell(4)));
				options[3] = (String.valueOf((Cell) currentRow.getCell(5)));
				quiz.setOption(options);
				quiz.setCrt((int) (Float.parseFloat(String.valueOf((Cell) currentRow.getCell(6))))-1);
				System.out.println(quiz);
				lstQuiz.add(quiz);
			}

			// Close WorkBook
			workbook.close();

			return lstQuiz;
		} catch (

		IOException e) {
			throw new RuntimeException("FAIL! -> message = " + e.getMessage());
		}
	}

	/**
	 * Convert Java Objects to JSON String
	 * 
	 * @param customers
	 * @param fileName
	 */
	private static String convertObjects2JsonString(List<Quiz> customers) {
		ObjectMapper mapper = new ObjectMapper();
		String jsonString = "";

		try {
			jsonString = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(customers);
		} catch (JsonProcessingException e) {
			e.printStackTrace();
		}

		return jsonString;
	}
}