package com.eleven.poi;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.Includes;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.PictureType;
import com.deepoove.poi.data.Pictures;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TableRenderData;
import com.deepoove.poi.data.Tables;
import com.eleven.transform.beans.ScreenShot;
import com.eleven.transform.utils.DateUtils;
import org.apache.commons.lang3.StringUtils;
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
        map.put("img", pictureRenderData);

        XWPFTable table = new XWPFTable();

        TableRenderData tableRenderData = new TableRenderData();
        List<RowRenderData> rows = new ArrayList();

        replaceText(map);
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
        XWPFTemplate compile = XWPFTemplate.compile("C:/Users/DELL/Desktop/transform/template.docx");
        compile.render(map);
        compile.writeToFile("C:/Users/DELL/Desktop/transform/template2.docx");
    }

    //根据实体对象 ，生成XWPFDocument
    public static XWPFDocument exportDataInfoWord(List<ScreenShot> list) throws NoSuchFieldException,IllegalAccessException {
        MyXWPFDocument doc = new MyXWPFDocument();
        XWPFTable table = doc.createTable(list.size() + 1, 12);
        List<TSType> types = ResourceUtil.getCacheTypes("primaryUse".toLowerCase());

        for(int colsIndex=0;colsIndex<fieldsNames.length;colsIndex++){
            XWPFTableCell cell = table.getRow(0).getCell(colsIndex);
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            XWPFParagraph p1 = cell.getParagraphs().get(0);
            XWPFRun r1 = p1.createRun();
            r1.setBold(true);
            r1.setText(fieldsNames[colsIndex]);
        }

        for(int rowIndex =1, listIndex =0; listIndex<list.size();rowIndex++,listIndex++){
            Class entityClass = list.get(listIndex).getClass();
            for(int colsIndex=0;colsIndex<fields.length;colsIndex++){
                Field field = entityClass.getDeclaredField(fields[colsIndex]);
                field.setAccessible(true);
                Object value = field.get(list.get(listIndex));


                XWPFTableCell cell = table.getRow(rowIndex).getCell(colsIndex);

                if(value instanceof Date){
                    cell.setText(new SimpleDateFormat("yyyy-MM-dd").format((Date)value));
                }else if(colsIndex == 4){
                    cell.setText(list.get(listIndex).getFgTypeName());
                }else if(colsIndex == 5){
                    cell.setText(list.get(listIndex).getFgVarietiesName());
                }
                else if(colsIndex == 7){
                    for(TSType tsType:types){
                        if(tsType.getTypecode().equals(list.get(listIndex).getPrimaryUse())){
                            cell.setText(tsType.getTypename());
                        }
                    }
                }
                else if(value instanceof Float){
                    cell.setText(String.valueOf(value));
                } else if (colsIndex == 11 && value instanceof String && value!=null) {
                    setCellImage(cell,value.toString());
                }else {
                    if(value!=null){
                        cell.setText(value.toString());
                    }
                }
            }
        }


        return doc;
    }

    //单元格写入图片
    private static  void setCellImage(XWPFTableCell cell,String urls)  {
        if(StringUtils.isBlank(urls))
            return;
        String [] urlArray = urls.split(",");
        String ctxPath=ResourceUtil.getConfigByName("webUploadpath");
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        XWPFParagraph newPara = paragraphs.get(0);
        XWPFRun imageCellRunn = newPara.createRun();
        for(String url:urlArray){
            String downLoadPath = ctxPath+File.separator + url;
            File image = new File(downLoadPath);
            if(!image.exists()){
                continue;
            }

            int format;
            if (url.endsWith(".emf")) {
                format = XWPFDocument.PICTURE_TYPE_EMF;
            } else if (url.endsWith(".wmf")) {
                format = XWPFDocument.PICTURE_TYPE_WMF;
            } else if (url.endsWith(".pict")) {
                format = XWPFDocument.PICTURE_TYPE_PICT;
            } else if (url.endsWith(".jpeg") || url.endsWith(".jpg")) {
                format = XWPFDocument.PICTURE_TYPE_JPEG;
            } else if (url.endsWith(".png")) {
                format = XWPFDocument.PICTURE_TYPE_PNG;
            } else if (url.endsWith(".dib")) {
                format = XWPFDocument.PICTURE_TYPE_DIB;
            } else if (url.endsWith(".gif")) {
                format = XWPFDocument.PICTURE_TYPE_GIF;
            } else if (url.endsWith(".tiff")) {
                format = XWPFDocument.PICTURE_TYPE_TIFF;
            } else if (url.endsWith(".eps")) {
                format = XWPFDocument.PICTURE_TYPE_EPS;
            } else if (url.endsWith(".bmp")) {
                format = XWPFDocument.PICTURE_TYPE_BMP;
            } else if (url.endsWith(".wpg")) {
                format = XWPFDocument.PICTURE_TYPE_WPG;
            } else {
                logger.error("Unsupported picture: " + url +
                        ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
                continue;
            }

            try (FileInputStream is = new FileInputStream(downLoadPath)) {
                imageCellRunn.addPicture(is, format, image.getName(), Units.toEMU(100), Units.toEMU(100)); // 200x200 pixels
            }catch (Exception e){
                logger.error(e.getMessage());
                e.printStackTrace();
            }
//            imageCellRunn.addBreak();

        }


    }
}
