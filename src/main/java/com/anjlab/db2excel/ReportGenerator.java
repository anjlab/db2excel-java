package com.anjlab.db2excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

import org.json.JSONException;
import org.json.JSONObject;

public class ReportGenerator {

    private RequestContext context;
    
    public ReportGenerator(RequestContext context) {
        this.context = context;
    }
    
    public static ReportGenerator createFromJsonRequest(String requestJson) throws JSONException {
        RequestContext context = RequestContext.fromJson(new JSONObject(requestJson));
        return new ReportGenerator(context);
    }

    public static ReportGenerator createFromJsonRequest(InputStream requestJson) throws UnsupportedEncodingException, IOException, JSONException {
        return createFromJsonRequest(IOUtils.readToEnd(requestJson, "UTF-8"));
    }

    public RequestContext getContext() {
        return context;
    }
    
    public void execute() throws ClassNotFoundException, SQLException, IOException {
        DataWriter writer = new ExcelWriter(context);
        
        Class.forName(context.getJdbcDriver());
        Connection connection = DriverManager.getConnection(context.getConnectionUrl());
        PreparedStatement statement = null;
        ResultSet resultSet = null;
        
        try {
            statement = connection.prepareStatement(context.getQuery());
            
            resultSet = statement.executeQuery();
            
            ResultSetMetaData metaData = resultSet.getMetaData();
            
            String[] headers = new String[metaData.getColumnCount()];
            for (int i = 0; i < headers.length; i++) {
                headers[i] = metaData.getColumnLabel(i + 1);
            }
            
            writer.writeHeaders(headers);
            
            Object[] values = new Object[headers.length];
            while (resultSet.next()) {
                for (int i = 0; i < values.length; i++) {
                    values[i] = resultSet.getObject(i + 1);
                }
                
                writer.writeRowValues(values);
            }
            
        } finally {
            if (resultSet != null) {
                try { resultSet.close(); } catch (SQLException e) { }
            }
            if (statement != null) {
                try { statement.close(); } catch (SQLException e) { }
            }
            if (connection != null) {
                try { connection.close(); } catch (SQLException e) { }
            }
            
            writer.close();
        }
    }
}
