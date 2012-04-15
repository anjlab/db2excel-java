package com.anjlab.db2excel;

import org.json.JSONException;
import org.json.JSONObject;

public class RequestContext {

    private String jdbcDriver;
    private String connectionUrl;
    private String query;
    private String templateFilePath;
    private String outputFilePath;
    private int dataSheetIndex;
    
    public static RequestContext fromJson(JSONObject json) throws JSONException {
        RequestContext context = new RequestContext();
        context.jdbcDriver = json.getString("jdbcDriver");
        context.connectionUrl = json.getString("connectionUrl");
        context.query = json.getString("query");
        context.templateFilePath = readString(json, "templateFilePath", null);
        context.outputFilePath = json.getString("outputFilePath");
        context.dataSheetIndex = readInt(json, "dataSheetIndex", 0);
        return context;
    }
    
    private static String readString(JSONObject json, String name, String defaultValue) throws JSONException {
        return json.has(name) ? json.getString(name) : defaultValue;
    }
    
    private static int readInt(JSONObject json, String name, int defaultValue) throws JSONException {
        return json.has(name) ? json.getInt(name) : defaultValue;
    }
    
    public String getJdbcDriver() {
        return jdbcDriver;
    }

    public String getConnectionUrl() {
        return connectionUrl;
    }

    public String getQuery() {
        return query;
    }

    public String getTemplateFilePath() {
        return templateFilePath;
    }

    public int getDataSheetIndex() {
        return dataSheetIndex;
    }

    public String getOutputFilePath() {
        return outputFilePath;
    }

}
