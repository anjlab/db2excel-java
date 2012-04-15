package com.anjlab.db2excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter implements DataWriter {

    private RequestContext context;
    
    private Workbook workbook;
    private Sheet sheet;
    
    private int rowIndex;
    
    public ExcelWriter(RequestContext context) throws IOException {
        this.context = context;
        
        if (context.getTemplateFilePath() != null) {
            File templateFile = new File(context.getTemplateFilePath());
            InputStream template = new FileInputStream(templateFile);
            workbook = templateFile.getName().toLowerCase().endsWith(".xls")
                     ? new HSSFWorkbook(template) 
                     : new XSSFWorkbook(template);
            template.close();
            sheet = workbook.getSheetAt(context.getDataSheetIndex());
        } else {
            workbook = context.getOutputFilePath().toLowerCase().endsWith(".xls")
                     ? new HSSFWorkbook()
                     : new XSSFWorkbook();
            sheet = workbook.createSheet();
        }
        
        rowIndex = 0;
    }

    @Override
    public void writeHeaders(String[] headers) {
        writeRow(headers);
    }

    private void writeRow(Object[] values) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        rowIndex++;
        
        Cell cell = null;

        for (int cellIndex = 0; cellIndex < values.length ; cellIndex++) {
            cell = row.getCell(cellIndex);
            if (cell == null) {
                cell = row.createCell(cellIndex);
            }
            
            Object value = values[cellIndex];
            
            if (value == null) continue;
            
            if (value instanceof Date) {
                cell.setCellValue((Date) value);
            } else if (value instanceof Number) {
                cell.setCellValue(((Number) value).doubleValue());
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else {
                cell.setCellValue(value.toString());
            }
        }
    }

    @Override
    public void writeRowValues(Object[] values) {
        writeRow(values);
    }

    @Override
    public void close() throws IOException {
        OutputStream out = new FileOutputStream(context.getOutputFilePath());
        workbook.write(out);
        out.close();
    }
}
