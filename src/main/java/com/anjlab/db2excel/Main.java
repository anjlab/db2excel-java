package com.anjlab.db2excel;


public class Main 
{
    public static void main( String[] args ) throws Exception
    {
        ReportGenerator generator = ReportGenerator.createFromJsonRequest(System.in);
        generator.execute();
    }
}
