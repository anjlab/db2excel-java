package com.anjlab.db2excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;

public class IOUtils {

    public static String readToEnd(InputStream is, String encoding)
            throws IOException, UnsupportedEncodingException {
          ByteArrayOutputStream baos = new ByteArrayOutputStream();
          
          byte[] buf = new byte[4096];
          int count;
          while ((count = is.read(buf)) != -1) {
            baos.write(buf, 0, count);
          }
          
          return new String(baos.toByteArray(), encoding);
    }

}
