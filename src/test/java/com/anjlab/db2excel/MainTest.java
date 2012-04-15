package com.anjlab.db2excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import junit.framework.Assert;

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
}
