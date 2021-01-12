package com.eleven.poi;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.render.WhereDelegate;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;

public class ReportImageRenderPolicy extends AbstractRenderPolicy<String> {

    @Override
    public void doRender(RenderContext<String> context) throws Exception {
        // anywhere delegate
        WhereDelegate where = context.getWhereDelegate();
        // any thing
        String thing = context.getThing();
        // do 图片
        InputStream inputStream = null;
        try {
            inputStream = getImageStream(thing);
            where.addPicture(inputStream, getPictureType(thing), 40, 30);
        } finally {
            IOUtils.closeQuietly(inputStream);
        }
        // clear
        clearPlaceholder(context, false);
    }

    private static int getPictureType(String url) {
        if (url.endsWith(".emf")) {
            return XWPFDocument.PICTURE_TYPE_EMF;
        } else if (url.endsWith(".wmf")) {
            return XWPFDocument.PICTURE_TYPE_WMF;
        } else if (url.endsWith(".pict")) {
            return XWPFDocument.PICTURE_TYPE_PICT;
        } else if (url.endsWith(".jpeg") || url.endsWith(".jpg")) {
            return XWPFDocument.PICTURE_TYPE_JPEG;
        } else if (url.endsWith(".png")) {
            return XWPFDocument.PICTURE_TYPE_PNG;
        } else if (url.endsWith(".dib")) {
            return XWPFDocument.PICTURE_TYPE_DIB;
        } else if (url.endsWith(".gif")) {
            return XWPFDocument.PICTURE_TYPE_GIF;
        } else if (url.endsWith(".tiff")) {
            return XWPFDocument.PICTURE_TYPE_TIFF;
        } else if (url.endsWith(".eps")) {
            return XWPFDocument.PICTURE_TYPE_EPS;
        } else if (url.endsWith(".bmp")) {
            return XWPFDocument.PICTURE_TYPE_BMP;
        } else if (url.endsWith(".wpg")) {
            return XWPFDocument.PICTURE_TYPE_WPG;
        } else {
            System.out.println("异常");
        }
        return 0;
    }

    /**
     * 获取网络图片流
     *
     * @param url
     * @return
     */
    public static InputStream getImageStream(String url) {
        try {
            HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
            connection.setReadTimeout(5000);
            connection.setConnectTimeout(5000);
            connection.setRequestMethod("GET");
            if (connection.getResponseCode() == HttpURLConnection.HTTP_OK) {
                InputStream inputStream = connection.getInputStream();
                return inputStream;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
