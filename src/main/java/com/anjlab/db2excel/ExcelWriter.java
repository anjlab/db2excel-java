package com.anjlab.db2excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelWriter implements DataWriter {

    private RequestContext context;
    
    private HSSFWorkbook workbook;
    private HSSFSheet sheet;
    
    private int rowIndex;
    
    public ExcelWriter(RequestContext context) throws IOException {
        this.context = context;
        
        if (context.getTemplateFilePath() != null) {
            File templateFile = new File(context.getTemplateFilePath());
            InputStream template = new FileInputStream(templateFile);
            workbook = new HSSFWorkbook(template);
            template.close();
            sheet = workbook.getSheetAt(context.getDataSheetIndex());
        } else {
            workbook = new HSSFWorkbook();
            sheet = workbook.createSheet();
        }
        
        rowIndex = 0;
    }

    @Override
    public void writeHeaders(String[] headers) {
        writeRow(headers);
    }

    private void writeRow(Object[] values) {
        HSSFRow row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        rowIndex++;
        
        HSSFCell cell = null;

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
