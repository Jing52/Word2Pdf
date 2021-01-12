package com.eleven.poi;

import com.deepoove.poi.data.*;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import com.eleven.transform.beans.ScreenShot;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

public class ReportScreenshotRenderPolicy extends AbstractRenderPolicy<ArrayList> {

    @Override
    protected void afterRender(RenderContext<ArrayList> context) {
        // 清空标签
        clearPlaceholder(context, true);
    }

    @Override
    public void doRender(RenderContext<ArrayList> context) throws Exception {
        XWPFRun run = context.getRun();
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);

        List<ScreenShot> screenShots = context.getThing();

        int num = screenShots.size();
        // 定义行列
        int row = (num % 2) != 0 ? (num / 2 + 1) * 2 : (num / 2) * 2;
        int col = 2;

        // 插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);
        table.removeBorders();

        for (int i = 0; i < row; i = i + 2) {
            for (int j = 0; j < col; j++) {
                //图片处理
                InputStream stream = getImageStream(screenShots.get(i + j).getUrl());
                XWPFParagraph pic = table.getRow(i).getCell(j).addParagraph();
                XWPFRun r1 = pic.createRun();
                r1.addPicture(stream, XWPFDocument.PICTURE_TYPE_PNG, "Generated", Units.toEMU(150), Units.toEMU(150));

                //文字处理
                //在表格中的行添加段落
                XWPFParagraph xwpfParagraph = table.getRow(i + 1).getCell(j).addParagraph();
                //获取光标的位置
                XmlCursor cursor = xwpfParagraph.getCTP().newCursor();
                xwpfParagraph.setAlignment(ParagraphAlignment.CENTER);
                xwpfParagraph.setVerticalAlignment(TextAlignment.CENTER);
                //在光标处插入表格
                XWPFTable xwpfTable = table.getRow(i + 1).getCell(j).insertNewTbl(cursor);
                xwpfTable.setTableAlignment(TableRowAlign.CENTER);
                //在表格中的表格中创建行
                XWPFTableRow xwpfTableRow = xwpfTable.createRow();
                //创建列
                XWPFTableCell xwpfTableCell = xwpfTableRow.createCell();
                xwpfTableCell.setWidth("2000");
                //在列中创建段落
                XWPFParagraph c1 = xwpfTableCell.addParagraph();
                XWPFRun rt1 = c1.createRun();
                rt1.setText(screenShots.get(i + j).getDesc());
                rt1.setBold(false);
                //移除段落，不然会出现一个空行
                xwpfTableCell.removeParagraph(0);
                if ((num % 2) != 0 && i == row - 2) {
                    break;
                }
            }
        }
    }

    private static PictureRenderData generateImage(String url, int width, int height) {
        PictureType pictureType = null;

        if (url.endsWith(".emf")) {
            pictureType = PictureType.EMF;
        } else if (url.endsWith(".wmf")) {
            pictureType = PictureType.WMF;
        } else if (url.endsWith(".pict")) {
            pictureType = PictureType.PICT;
        } else if (url.endsWith(".jpeg") || url.endsWith(".jpg")) {
            pictureType = PictureType.JPEG;
        } else if (url.endsWith(".png")) {
            pictureType = PictureType.PNG;
        } else if (url.endsWith(".dib")) {
            pictureType = PictureType.DIB;
        } else if (url.endsWith(".gif")) {
            pictureType = PictureType.GIF;
        } else if (url.endsWith(".tiff")) {
            pictureType = PictureType.TIFF;
        } else if (url.endsWith(".eps")) {
            pictureType = PictureType.EPS;
        } else if (url.endsWith(".bmp")) {
            pictureType = PictureType.BMP;
        } else if (url.endsWith(".wpg")) {
            pictureType = PictureType.WPG;
        } else {
            System.out.println("异常");
        }
        return Pictures.ofUrl(url, pictureType).size(width, height).create();
    }

    /**
     * 获取网络图片流
     *
     * @param url
     * @return
     */
    private static InputStream getImageStream(String url) {
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
