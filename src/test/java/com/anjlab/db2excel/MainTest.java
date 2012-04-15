package com.anjlab.db2excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import junit.framework.Assert;

import org.json.JSONObject;
import org.junit.Test;

public class MainTest {

    @Test
    public void testGenerator() throws Exception {
        File file = new File("src/test/resources/request.json");
        InputStream input = new FileInputStream(file);
        ReportGenerator generator = ReportGenerator.createFromJsonRequest(input);
        input.close();
        
        File out = new File(generator.getContext().getOutputFilePath());
        if (out.exists()) {
            out.delete();
        }
        
        generator.execute();
        Assert.assertTrue(out.exists());
    }
    
    @Test
    public void testGeneratorWithTemplate() throws Exception {
        File file = new File("src/test/resources/request_with_template.json");
        InputStream input = new FileInputStream(file);
        ReportGenerator generator = ReportGenerator.createFromJsonRequest(input);
        input.close();
        
        File out = new File(generator.getContext().getOutputFilePath());
        if (out.exists()) {
            out.delete();
        }
        
        generator.execute();
        Assert.assertTrue(out.exists());
    }
    
    @Test
    public void testMaxRowNumXLS() throws Exception {
        String json = "{'jdbcDriver':'','connectionUrl':'','query':'','outputFilePath':'target/maxrownum.xls'," +
                        "'templateFilePath':'src/test/resources/result_template.xls'}";
        ExcelWriter writer = new ExcelWriter(RequestContext.fromJson(new JSONObject(json)));
        Object[] values = new Object[] { "test" };
        try {
            for (int i = 0; i < 65536 + 1; i++) {
                writer.writeRowValues(values);
            }
            Assert.fail("HSSFRow supports maximum 65536 rows");
        } catch (IllegalArgumentException e) {
            Assert.assertEquals("Invalid row number (65536) outside allowable range (0..65535)", e.getMessage());
        }
    }
    
    @Test
    public void testMaxRowNumXSLX() throws Exception {
        String json = "{'jdbcDriver':'','connectionUrl':'','query':'','outputFilePath':'target/maxrownum.xlsx'," +
                        "'templateFilePath':'src/test/resources/result_template.xlsx'}";
        ExcelWriter writer = new ExcelWriter(RequestContext.fromJson(new JSONObject(json)));
        Object[] values = new Object[] { "test" };
        for (int i = 0; i < 65536 + 1; i++) {
            writer.writeRowValues(values);
        }
    }
}
