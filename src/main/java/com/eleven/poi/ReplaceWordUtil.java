package com.eleven.poi;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.data.*;
import com.deepoove.poi.data.style.TableStyle;
import com.eleven.transform.beans.ScreenShot;
import com.eleven.transform.utils.DateUtils;
import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static javax.swing.UIManager.put;

/**
 * @Author: Evan
 * @CreateTime: 2021-01-08
 * @Description:
 */
public class ReplaceWordUtil {
    public static void main(String[] args) throws IOException {
        Map<String, Object> map = new HashMap<>();
        map.put("diagnoseNo", "HZ520520");
        map.put("name", "张三");
        map.put("sex", "男");
        map.put("age", "56");
        map.put("department", "精神科");
        map.put("bedNo", "1001");
        map.put("hospNo", "2001001");
        map.put("opc", "30001001");
        map.put("doctor", "范进");
        map.put("sendDate", "2020/12/01");
        map.put("diagnose", "1送检为福尔马林固定标本，包装外附患者身份信息，标注送检组织取样部位为“嘴唇肿块”：皮肤及皮下组织一块，大小1.2×0.4×0.3cm，皮肤表面见一灰褐色区域，直径0.2cm，切面灰白质中，对剖全取。");
        map.put("view", "2送检为福尔马林固定标本，包装外附患者身份信息，标注送检组织取样部位为“嘴唇肿块”：皮肤及皮下组织一块，大小1.2×0.4×0.3cm，皮肤表面见一灰褐色区域，直径0.2cm，切面灰白质中，对剖全取。");
        map.put("currentDiagnose", "符合");
        map.put("expertName", "郭德湖");
        map.put("recheckName", "王得金");
        map.put("reportDate", DateUtils.getCurrentDate());
        PictureRenderData pictureRenderData = generateImage("https://pcp.oss.histomed.com/report/image/20210107/2100130-1610005479145-175936.png", 150, 150);
        map.put("expertImg", "https://pcp.oss.test.hengdaomed.com/expert/sign/20201228/3-1609134906364-356818.jpg");
        map.put("recheckImg", "https://pcp.oss.test.hengdaomed.com/expert/sign/20201228/2-1609134880739-239864.jpg");

        List<ScreenShot> screenShots = new ArrayList<>();
        ScreenShot screenShot = new ScreenShot();
        screenShot.setDesc("1111111111");
        screenShot.setUrl("https://pcp.oss.histomed.com/report/image/20210107/2100130-1610005410257-774571.png");
        screenShots.add(screenShot);

        screenShot = new ScreenShot();
        screenShot.setDesc("2222222222");
        screenShot.setUrl("https://pcp.oss.histomed.com/report/image/20210107/2100130-1610005422120-346985.png");
        screenShots.add(screenShot);

        screenShot = new ScreenShot();
        screenShot.setDesc("3333333333");
        screenShot.setUrl("https://pcp.oss.histomed.com/report/image/20210107/2100130-1610005436155-671268.png");
        screenShots.add(screenShot);

        screenShot = new ScreenShot();
        screenShot.setDesc("4444444444");
        screenShot.setUrl("https://pcp.oss.histomed.com/report/image/20210107/2100130-1610005479145-175936.png");
        screenShots.add(screenShot);

        screenShot = new ScreenShot();
        screenShot.setDesc("5555555555");
        screenShot.setUrl("https://pcp.oss.histomed.com/report/image/20210105/TIS73034009-1609831299497-871738.png");
        screenShots.add(screenShot);
        map.put("table", screenShots);

//        replaceText(map);

        ConfigureBuilder builder = Configure.builder();
        Configure config = builder.bind("table", new ReportScreenshotRenderPolicy())
                .bind("expertImg", new ReportImageRenderPolicy())
                .bind("recheckImg", new ReportImageRenderPolicy())
                .build();

        //生成替换数据后的word
        XWPFTemplate.compile("/Users/Evan/Desktop/transform/template.docx", config).render(map).writeToFile("/Users/Evan/Desktop/transform/template2.docx");

        //word转换成pdf
        word2pdf("/Users/Evan/Desktop/transform/template2.docx", "HZ520520", "/Users/Evan/Desktop/transform/");
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

    private static void replaceText(Map<String,Object> map) throws IOException {
        //核心API采用了极简设计，只需要一行代码
        XWPFTemplate compile = XWPFTemplate.compile("/Users/Evan/Desktop/transform/template.docx");
        compile.render(map);
        compile.writeToFile("/Users/Evan/Desktop/transform/template1.docx");
    }

    /**
     * word转pdf
     *
     * @param sourcePath
     * @return
     */
    public static String word2pdf(String sourcePath, String diagnoseNo, String pdfPath) {
        Document document = new Document(sourcePath);
        String fileName = diagnoseNo + "-" + DateUtils.getTimeStamp() + "-" + "000000" + ".pdf";
        String pdfNoMarkPath = pdfPath + fileName;
        document.saveToFile(pdfNoMarkPath, FileFormat.PDF);
        return fileName;
    }
}
