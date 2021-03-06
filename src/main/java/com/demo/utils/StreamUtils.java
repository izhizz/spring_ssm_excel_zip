package com.demo.utils;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URL;
import java.net.URLEncoder;

/**
 * @author fengzhi
 */
public class StreamUtils {
    private static byte[] buffer = new byte[1024];
//==========================================file===========================================================
    /**
     * 下载外部资源通过流存储到磁盘上
     *
     * @param file         外部资源路径
     * @param savePath     存储文件夹
     * @param filePathName 存储文件路径名 ;例：c:/aa/a.doc
     * @throws IOException
     */
    public static void writeDisk(String file, String savePath, String filePathName) throws IOException {
        URL url = new URL(file);
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        InputStream fis = url.openConnection().getInputStream();
        int r = 0;
        while ((r = fis.read(buffer)) != -1) {
            output.write(buffer, 0, r);
        }
        File dir = new File(savePath);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        FileOutputStream fileOutputStream = new FileOutputStream(new File(filePathName));
        fileOutputStream.write(output.toByteArray());
        fileOutputStream.close();
        fis.close();
    }

    /**
     * 返回客户端响响应流
     *
     * @param response
     * @param in
     * @throws IOException
     */
    public static void responseClient(HttpServletResponse response, InputStream in) throws IOException {
        OutputStream out = response.getOutputStream();
        int b;
        while ((b = in.read()) != -1) {
            out.write(b);
        }
        in.close();
        out.close();
    }

    /**
     * 设置下载响应头
     *
     * @param name
     * @param response
     * @throws UnsupportedEncodingException
     */
    public static void reponseStreamUtf8(String name, HttpServletResponse response) throws UnsupportedEncodingException {
        //转换中文否则可能会产生乱码
        name = URLEncoder.encode(name, "UTF-8");
        // 指明response的返回对象是文件流
        response.setContentType("application/octet-stream");
        // 设置在下载框默认显示的文件名
        response.setHeader("Content-Disposition", "attachment;filename=" + name);
    }
}
