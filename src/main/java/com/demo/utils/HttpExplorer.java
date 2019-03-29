package com.demo.utils;

import org.springframework.util.Base64Utils;

import javax.servlet.http.HttpServletRequest;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

public class HttpExplorer {
    public static String getClientExplorerType(HttpServletRequest request) {
        String agent = request.getHeader("USER-AGENT");
        if (agent != null && agent.toLowerCase().indexOf("firefox") > 0) {
            return "firefox";
        } else if (agent != null && agent.toLowerCase().indexOf("msie") > 0) {
            return "ie";
        } else if (agent != null && agent.toLowerCase().indexOf("chrome") > 0) {
            return "chrome";
        } else if (agent != null && agent.toLowerCase().indexOf("opera") > 0) {
            return "opera";
        } else if (agent != null && agent.toLowerCase().indexOf("safari") > 0) {
            return "safari";
        }
        return "others";
    }

    public static String getFileNameEncoder(HttpServletRequest request, String fileName) throws UnsupportedEncodingException {
        if ("firefox".equals(getClientExplorerType(request))) {
            //火狐浏览器自己会对URL进行一次URL转码所以区别处理
            return "=?UTF-8?B?" + (new String(Base64Utils.encodeToString(fileName.getBytes("UTF-8")))) + "?=";
        } else {
            return URLEncoder.encode(fileName, "UTF-8");
        }
    }
}
