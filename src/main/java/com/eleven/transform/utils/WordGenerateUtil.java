package com.eleven.transform.utils;

import com.eleven.transform.beans.ReportParam;
import com.eleven.transform.beans.ScreenShot;
import com.spire.doc.AutoFitBehaviorType;
import com.spire.doc.Document;
import com.spire.doc.DocumentProperty;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.Table;
import com.spire.doc.TableRow;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.TextSelection;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.TextRange;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.List;

/**
 * @Author: Evan
 * @CreateTime: 2021-01-05
 * @Description:
 */
public class WordGenerateUtil {

    public static void replace(String sourcePath, ReportParam params, String pdfPath) throws IOException {
        //加载Word文档
        Document document = new Document(sourcePath);
        //添加节
        Section section = document.addSection();

        //使用新文本替换文档中的指定文本
        replaceText(document, params);

        replaceImage(document, "${header}", "C:/Users/DELL/Desktop/tramsform/header.jpg", 400, 50);
        replaceImage(document, "${footer}", "C:/Users/DELL/Desktop/tramsform/footer.jpg", 352, 30);

        //添加大体所见
        addView(section, params.getView());
        //添加镜像所见
        addMicro(section, params.getShots());
        //添加诊断意见
        addDiagnose(section, params.getDiagnose());

//        //用图片替换 Word 文档中的特定文本
//        addImages2Excel(document, params.getShots());

        //保存文档
//        document.saveToFile("C:/Users/DELL/Desktop/tramsform/test1.docx", FileFormat.Docx_2013);

        //word转pdf
        word2pdf(document, params.getDiagnoseNo(), pdfPath);
    }

    /**
     * 大体所见
     *
     * @param section
     * @return
     */
    public static void addView(Section section, String view) {
        //添加第一个段落
        Paragraph paragraph1 = section.addParagraph();
        //添加图片到段落
        paragraph1.appendText("大体所见:");
        //添加第二个段落
        section.addParagraph();
        //添加第三个段落
        Paragraph paragraph3 = section.addParagraph();
        //添加图片到段落
        paragraph3.appendText(view);
    }

    /**
     * 镜下所见
     *
     * @param section
     * @param shots
     * @return
     */
    public static void addMicro(Section section, List<ScreenShot> shots) throws IOException {
        //添加第一个段落
        Paragraph paragraph1 = section.addParagraph();
        //添加图片到段落
        paragraph1.appendText("镜下所见:");
        //添加第二个段落
        section.addParagraph();
        //添加图片到段落
        addImages2Excel(section, shots);
    }

    /**
     * 诊断意见
     *
     * @param section
     * @param diagnose
     * @return
     */
    public static void addDiagnose(Section section, String diagnose) {
        //添加第一个段落
        Paragraph paragraph1 = section.addParagraph();
        //添加文本到段落
        paragraph1.appendText("诊断意见:");
        //添加第二个段落
        section.addParagraph();
        //添加第二个段落
        Paragraph paragraph3 = section.addParagraph();
        //添加文本到段落
        paragraph3.appendText(diagnose);
    }

    /**
     * 替换文本
     *
     * @param document
     */
    private static void replaceText(Document document, ReportParam params) {
        document.replace("${name}", params.getPtName(), false, true);
        document.replace("${age}", params.getPtAge(), false, true);
        document.replace("${gender}", params.getPtGender(), false, true);
        document.replace("${diagnosisNo}", params.getDiagnoseNo(), false, true);
        document.replace("${department}", params.getDepartment(), false, true);
        document.replace("${bedNo}", params.getBedNo(), false, true);
        document.replace("${hospNo}", params.getHospNo(), false, true);
        document.replace("${doctor}", params.getDoctor(), false, true);
        document.replace("${date}", params.getSendDate(), false, true);
        document.replace("${site}", params.getSite(), false, true);
        document.replace("${diagnose}", params.getDiagnose(), false, true);
        document.replace("${view}", params.getView(), false, true);
        document.replace("${currentDiagnose}", params.getCurrentDiagnose(), false, true);
    }

    /**
     * 替换图片
     *
     * @param document
     */
    private static void replaceImage(Document document, String transText, String imgPath, float width, float height) {
        //查找文档中的字符串transText
        TextSelection[] selections = document.findAllString(transText, true, true);

        //用图片替换文字
        int index = 0;
        TextRange range = null;
        for (Object obj : selections) {
            TextSelection textSelection = (TextSelection) obj;
            DocPicture pic = new DocPicture(document);
            pic.loadImage(imgPath);
            //设置图片宽度
            pic.setWidth(width);
            //设置图片高度
            pic.setHeight(height);
            range = textSelection.getAsOneRange();
            index = range.getOwnerParagraph().getChildObjects().indexOf(range);
            range.getOwnerParagraph().getChildObjects().insert(index, pic);
            range.getOwnerParagraph().getChildObjects().remove(range);
        }
    }

    /**
     * 添加图片
     *
     * @param document
     */
    private static void addImage(Document document) {
        //添加节
        Section section = document.addSection();
        //添加第一个段落
        Paragraph paragraph1 = section.addParagraph();
        //添加图片到段落
        DocPicture picture = paragraph1.appendPicture("C:\\Users\\DELL\\Desktop\\生产图片\\0.jpg");
        //设置图片宽度
        picture.setWidth(200f);
        //设置图片高度
        picture.setHeight(150f);
    }

    /**
     * 添加图片到 Word 表格
     *
     * @param section
     * @return
     */
    public static void addImages2Excel(Section section, List<ScreenShot> screenShots) throws IOException {

        int cells = screenShots.size();
        int rows = (cells % 2) != 0 ? (cells / 2 + 1) * 2 : (cells / 2) * 2;
        int columns = 2;
        //添加表格
        Table table = section.addTable(true);
        table.resetCells(rows, columns);

        //添加图片到单元格，并自定义图片大小
        //添加图片到单元格（0，0）
        for (int i = 0; i < rows; i = i + 2) {
            for (int j = 0; j < columns; j++) {
                InputStream imageStream = getImageStream(screenShots.get(i + j).getUrl());
                DocPicture picture = table.getRows().get(i).getCells().get(j).addParagraph().appendPicture(imageStream);
                System.out.println("i:" + i + "j:" + j);
                TableRow dataRow = table.getRows().get(i + 1);
                dataRow.getCells().get(j).addParagraph().appendText(screenShots.get(i + j).getDesc());
                picture.setWidth(100f);
                picture.setHeight(100f);
                if ((cells % 2) != 0 && i == rows - 2) {
                    break;
                }
            }
        }

        //设置表格大小自适应内容
        table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
    }

    /**
     * 添加图片
     *
     * @param document
     * @return
     */
    public static String word2pdf(Document document, String diagnoseNo, String pdfPath) {
        String fileName = diagnoseNo + "-" + DateUtils.getTimeStamp() + "-" + "000000" + ".pdf";
        String pdfNoMarkPath = pdfPath + fileName;
        document.saveToFile(pdfNoMarkPath, FileFormat.PDF);
        return fileName;
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
            System.out.println("获取网络图片出现异常，图片路径为：" + url);
            e.printStackTrace();
        }
        return null;
    }
}
