package com.example.demo.util;


import java.io.IOException;
import java.util.List;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import com.example.demo.model.CategoriaProdotto;
import com.example.demo.model.Excel;
 
@Component
public class ExcelExport {
 
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private List<Excel> prodotti;
	
	
	public ExcelExport(List<Excel> listExcel) {
		this.prodotti=listExcel;
		workbook = new XSSFWorkbook();
	}

	public void export(HttpServletResponse response) throws IOException{
		writeHeaderLine();
		writeDataLines();

		ServletOutputStream outputStream=response.getOutputStream();
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
	}

	private void writeHeaderLine() {
		sheet=workbook.createSheet("prodotti");
		sheet.setColumnWidth(0,3500);
		sheet.setColumnWidth(1,3500);
		sheet.setColumnWidth(2,3500);
		Row row = sheet.createRow(0);
		CellStyle style = workbook.createCellStyle();

		XSSFFont font=workbook.createFont();
		font.setBold(true);
		font.setFontHeight(16);
		style.setFont(font);
		style.setWrapText(true);
		style.setAlignment(HorizontalAlignment.CENTER);

		createCell(row,0,"Nome",style);
		createCell(row,1,"Categoria",style);
		createCell(row,2,"Prezzo",style);
	}

	private void writeDataLines() {
		int rowCount=1;
		CellStyle style=workbook.createCellStyle();
		XSSFFont font=workbook.createFont();
		font.setFontHeight(12);
		style.setFont(font);
		style.setWrapText(true);
		style.setAlignment(HorizontalAlignment.CENTER);

		for(Excel excel:prodotti) {
			Row row=sheet.createRow(rowCount++);
			int columnCount=0;
			createCell(row, columnCount++, excel.getNomeProdotto(), style);
			createCell(row, columnCount++, excel.getCategoriaProdotto(), style);
			createCell(row, columnCount++, excel.getPrezzo(), style);
		}
	}
	
	private void createCell(Row row,int columnCount, Object value,CellStyle style) {
		sheet.autoSizeColumn(columnCount);
		Cell cell=row.createCell(columnCount);
		if(value instanceof Long) {
			cell.setCellValue((Long) value);
		}else if(value instanceof Integer) {
			cell.setCellValue((Integer) value);
		}else if(value instanceof Double) {
			cell.setCellValue((Double) value);
		}else if(value instanceof String){
			cell.setCellValue((String) value);
		}else if(value instanceof CategoriaProdotto){
			cell.setCellValue(value.toString());
		}
		cell.setCellStyle(style);
	}

}
