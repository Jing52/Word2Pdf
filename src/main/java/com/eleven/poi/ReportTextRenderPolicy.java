package com.eleven.poi;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ReportTextRenderPolicy extends AbstractRenderPolicy<String> {

    @Override
    public void doRender(RenderContext<String> context) throws Exception {
        // anywhere
        XWPFRun where = context.getWhere();
        // anything
        String thing = context.getThing();
        // do 文本
        where.setText(thing, 0);
    }
}
