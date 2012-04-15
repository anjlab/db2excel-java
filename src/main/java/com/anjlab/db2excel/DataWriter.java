package com.anjlab.db2excel;

import java.io.IOException;

public interface DataWriter {

    void writeHeaders(String[] headers);

    void writeRowValues(Object[] values);
    
    void close() throws IOException;

}
