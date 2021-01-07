package com.eleven.transform;

import com.eleven.transform.beans.ReportParam;
import com.eleven.transform.beans.ScreenShot;
import com.eleven.transform.utils.WordGenerateUtil;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author: Evan
 * @CreateTime: 2021-01-05
 * @Description:
 */
public class Test {

    public static void main(String[] args) {
        ReportParam params = new ReportParam();
        params.setDiagnoseNo("HZ2020122001");
        params.setPtName("David");
        params.setPtGender("男");
        params.setPtAge("25");
        params.setDepartment("放射科");
        params.setBedNo("5201314");
        params.setHospNo("20201314");
        params.setSendDate("2020-12-01");
        params.setDoctor("李翰");
        params.setSite("肺");
        params.setView("1送检为福尔马林固定标本，包装外附患者身份信息，标注送检组织取样部位为“嘴唇肿块”：皮肤及皮下组织一块，大小1.2×0.4×0.3cm，皮肤表面见一灰褐色区域，直径0.2cm，切面灰白质中，对剖全取。");
        params.setDiagnose("2送检为福尔马林固定标本，包装外附患者身份信息，标注送检组织取样部位为“嘴唇肿块”：皮肤及皮下组织一块，大小1.2×0.4×0.3cm，皮肤表面见一灰褐色区域，直径0.2cm，切面灰白质中，对剖全取。");
        params.setCurrentDiagnose("符合");

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

        params.setShots(screenShots);

        try {
            WordGenerateUtil.replace("C:/Users/DELL/Desktop/tramsform/test.docx", params, "C:/Users/DELL/Desktop/tramsform/");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
