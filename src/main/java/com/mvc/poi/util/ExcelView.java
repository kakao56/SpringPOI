package com.mvc.poi.util;

import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mvc.poi.vo.EmpVO;

public class ExcelView extends AbstractExcelPOIView {

	@SuppressWarnings("unchecked")
	@Override
	protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request, HttpServletResponse response) throws Exception {
		
		List<EmpVO> empList = (List<EmpVO>) model.get("empList");

		// Sheet 생성
		Sheet sheet = workbook.createSheet("empList");
		Row row = null;
		Cell cell = null;
		int rowCount = 0;	
		int cellCount = 0;
				
		CellStyle headStyle = workbook.createCellStyle();		
	    headStyle.setBorderTop(BorderStyle.THIN);
	    headStyle.setBorderBottom(BorderStyle.THIN);
	    headStyle.setBorderLeft(BorderStyle.THIN);
	    headStyle.setBorderRight(BorderStyle.THIN);
	    headStyle.setFillForegroundColor(HSSFColorPredefined.GREY_25_PERCENT.getIndex());
	    headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    headStyle.setAlignment(HorizontalAlignment.CENTER);
	    
	    CellStyle strBodyStyle = workbook.createCellStyle();
	    strBodyStyle.setBorderTop(BorderStyle.THIN);
	    strBodyStyle.setBorderBottom(BorderStyle.THIN);
	    strBodyStyle.setBorderLeft(BorderStyle.THIN);
	    strBodyStyle.setBorderRight(BorderStyle.THIN);
	    
	    // head Cell 생성
	    row = sheet.createRow(rowCount++);	 
	    
	    cell = row.createCell(cellCount++);	    
	    cell.setCellStyle(headStyle);
	    cell.setCellValue("EMPNO");	
	    
	    cell = row.createCell(cellCount++);	    
	    cell.setCellStyle(headStyle);
	    cell.setCellValue("ENAME");		
	    
	    cell = row.createCell(cellCount++);	    
	    cell.setCellStyle(headStyle);
	    cell.setCellValue("JOB");		
	    
	    cell = row.createCell(cellCount++);	    
	    cell.setCellStyle(headStyle);
	    cell.setCellValue("MGR");		
	    
	    cell = row.createCell(cellCount++);	    
	    cell.setCellStyle(headStyle);
	    cell.setCellValue("HIREDATE");		
	    
	    cell = row.createCell(cellCount++);	    
	    cell.setCellStyle(headStyle);
	    cell.setCellValue("SAL");		
	    
	    cell = row.createCell(cellCount++);	    
	    cell.setCellStyle(headStyle);
	    cell.setCellValue("COMM");		
	    
	    cell = row.createCell(cellCount++);	    
	    cell.setCellStyle(headStyle);
	    cell.setCellValue("DEPTNO");	
	    
		for (int i = 0; i < cellCount; i++) {//head 가로폭 자동
			
			sheet.autoSizeColumn(i);
			sheet.setColumnWidth(i, (sheet.getColumnWidth(i)) + 512);
		
		}
		
		
		// 데이터 Cell 생성
		for (EmpVO emp : empList) {
			row = sheet.createRow(rowCount++);
			cellCount = 0;
			
			cell = row.createCell(cellCount++);
			cell.setCellValue(emp.getEMPNO()); 
			cell.setCellStyle(strBodyStyle);
			
			cell =  row.createCell(cellCount++);
			cell.setCellValue(emp.getENAME()); 
			cell.setCellStyle(strBodyStyle);
			
			cell = row.createCell(cellCount++);
			cell.setCellValue(emp.getJOB()); 
			cell.setCellStyle(strBodyStyle);
			
			cell = row.createCell(cellCount++);
			cell.setCellValue(emp.getMGR()); 
			cell.setCellStyle(strBodyStyle);
			
			cell = row.createCell(cellCount++);
			cell.setCellValue(emp.getHIREDATE()); 
			cell.setCellStyle(strBodyStyle);
			
			cell = row.createCell(cellCount++);
			cell.setCellValue(emp.getSAL()); 
			cell.setCellStyle(strBodyStyle);
			
			cell = row.createCell(cellCount++);
			cell.setCellValue(emp.getCOMM()); 
			cell.setCellStyle(strBodyStyle);
			
			cell = row.createCell(cellCount++);
			cell.setCellValue(emp.getDEPTNO()); 
			cell.setCellStyle(strBodyStyle);

		}

	}

	@Override
	protected Workbook createWorkbook() {
		return new XSSFWorkbook();
	}

}
