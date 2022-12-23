package org.example;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.finders.RangeFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.*;
import org.docx4j.wml.*;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Docx4jTest {

    static String TEMPLATE_DOCX_FILE_PATH = "\\模板.docx";
    static String OUT_DOCX_FILE_PATH = "\\模板${date}.docx";

    static String path = "\\测试HTML转Docx${date}.docx";

    private static String richText = "";

    public static void main(String[] args) throws Exception {
        Map<String, String> valueMap = new HashMap<>();
        valueMap.put("docNumber", "2022-00001");
        valueMap.put("fdCity", "广东深圳");
        valueMap.put("fdDate", "2022-12-19");
//        valueMap.put("fdContent", "");
        // 第一步, 将富文本添加进去
        String fileName = replaceContent(richText);
        // 第二步, 将其他文本替换进去
        replaceOther(valueMap, fileName);
    }

    private static void replaceOther(Map<String, String> valueMap, String fileName) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(fileName));
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
        mainDocumentPart.variableReplace(valueMap);
        wordMLPackage.save(new File(fileName));
    }

    private static String replaceContent(String richText) {
        // 拼接html格式内容
        StringBuffer sbf = new StringBuffer();
        // 这里拼接一下html标签,便于word文档能够识别
        sbf.append("<!DOCTYPE html PUBLIC \"-//W3C//DTD HTML 4.01 Strict//EN\" \"http://www.w3.org/TR/html4/strict.dtd\">");
        sbf.append("<html lang=\"zh\">");
        sbf.append("<head>");
        sbf.append("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />");
        sbf.append("</head>");
        sbf.append("<body>");
        // 富文本内容
        sbf.append(richText);
        sbf.append("</body></html>");
        return replaceTemplate(sbf.toString());
    }

    public static void test(String richContent) throws Exception {
//        richConvertToDocx(richContent);
//        testCreateDocx();
//        replaceTemplate(richContent);
    }

    private static void testCreateDocx() throws Docx4JException {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(TEMPLATE_DOCX_FILE_PATH));
        File file = new File("E:\\ALL\\g_工作\\x_项目\\w_万物云\\需求收集\\20221130_最终版\\模板\\test\\测试生成docx.docx");
        wordMLPackage.getMainDocumentPart().addParagraphOfText("你好!");
        wordMLPackage.save(file);
    }

    private static String replaceTemplate(String richContent) {
        WordprocessingMLPackage aPackage;
        try {
            aPackage = WordprocessingMLPackage.load(new File(TEMPLATE_DOCX_FILE_PATH));
            MainDocumentPart mainDocumentPart = aPackage.getMainDocumentPart();
            // 测试弹窗
//            mainDocumentPart.getContent().add(1, mainDocumentPart.getContent().get(0));



            // 方式一，将html追加至文档最后
//            mainDocumentPart.addAltChunk(AltChunkType.Html, buildHtml(richContent).getBytes(Charsets.UTF_8));
            // 方式二，将html插入至指定位置
            XHTMLImporterImpl importer = new XHTMLImporterImpl(aPackage);

            // 提取正文中所有段落
            List<Object> paragraphs = mainDocumentPart.getContent();
            // 提取书签并创建书签的游标
            RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
            new TraversalUtil(paragraphs, rt);

            // 遍历书签
            for (CTBookmark bm : rt.getStarts()) {
                // 这儿可以对单个书签进行操作，也可以用一个map对所有的书签进行处理
                if (bm.getName().equals("fdContent")) {
                    // 获取该书签的父级段落
                    P p = (P) (bm.getParent());
                    paragraphs.addAll(paragraphs.indexOf(p), importer.convert(richContent, null));
                    paragraphs.remove(p);
                }
            }
            //  将构建的word内容保存到临时文件
            String fileName = OUT_DOCX_FILE_PATH.replace("${date}", String.valueOf(System.currentTimeMillis()));
            aPackage.save(new File(fileName));
            return fileName;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }

    public static void docx4jHtmlConvertToDocx(String html) throws Exception {
        WordprocessingMLPackage pack = WordprocessingMLPackage.createPackage();
        NumberingDefinitionsPart numDefPart = new NumberingDefinitionsPart();
        pack.getMainDocumentPart().addTargetPart(numDefPart);
        numDefPart.unmarshalDefaultNumbering();
        // Convert XHTML + CSS to WordML content.
        // XHTML must be well formed XML.
        XHTMLImporterImpl importer = new XHTMLImporterImpl(pack);
        // Convert the well formed XHTML contained in file to a list of WML objects.
        List<Object> list = importer.convert(html, null);
        MainDocumentPart mainPart = pack.getMainDocumentPart();
        // Add all WML objects to MainDocumentPart.
        mainPart.getContent().addAll(list);
//        // create footer
//        Ftr footer = createFooterWithPageNumber();
//        FooterPart footerPart = new FooterPart();
//        footerPart.setPackage(pack);
//        footerPart.setJaxbElement(footer);
//        Relationship footerRelation = mainPart.addTargetPart(footerPart);
//        FooterReference footerRef = factory.createFooterReference();
//        footerRef.setId(footerRelation.getId());
//        footerRef.setType(HdrFtrRef.DEFAULT);
//        SectPr sectPr = pack.getDocumentModel().getSections().get(0).getSectPr();
//        sectPr.getEGHdrFtrReferences().add(footerRef);
//        // create header
//        Hdr header = createHeader("北京****科技有限公司");
//        HeaderPart headerPart = new HeaderPart();
//        headerPart.setPackage(pack);
//        headerPart.setJaxbElement(header);
//        Relationship headerRelation = mainPart.addTargetPart(headerPart);
//        HeaderReference headerRef = factory.createHeaderReference();
//        headerRef.setId(headerRelation.getId());
//        headerRef.setType(HdrFtrRef.DEFAULT);
//        sectPr.getEGHdrFtrReferences().add(headerRef);
        // set compatibilityMode to 15
        // to avoid Word 365/2016 saying "Compatibility Mode"
        DocumentSettingsPart settingsPart = mainPart.getDocumentSettingsPart(true);
        CTCompat compat = Context.getWmlObjectFactory().createCTCompat();
        compat.setCompatSetting("compatibilityMode", "http://schemas.microsoft.com/office/word", "15");
        settingsPart.getContents().setCompat(compat);
        // save to a file
        pack.save(new File(path.replace("${date}", String.valueOf(System.currentTimeMillis()))));
    }

    public static void replaceText(CTBookmark bm, Object object) throws Exception {
        if (object == null) {
            return;
        }
        if (bm.getName() == null){
            return;
        }
        String value = object.toString();
        try {
            List<Object> theList = null;
            ParaRPr rpr = null;
            if (bm.getParent() instanceof P) {
                PPr pprTemp = ((P) (bm.getParent())).getPPr();
                if (pprTemp == null) {
                    rpr = null;
                } else {
                    rpr = ((P) (bm.getParent())).getPPr().getRPr();
                }
                theList = ((ContentAccessor) (bm.getParent())).getContent();
            } else {
                return;
            }
            int rangeStart = -1;
            int rangeEnd = -1;
            int i = 0;
            for (Object ox : theList) {
                Object listEntry = XmlUtils.unwrap(ox);
                if (listEntry.equals(bm)) {
                    if (((CTBookmark) listEntry).getName() != null) {
                        rangeStart = i + 1;
                    }
                } else if (listEntry instanceof CTMarkupRange) {
                    if (((CTMarkupRange) listEntry).getId().equals(bm.getId())) {
                        rangeEnd = i - 1;
                        break;
                    }
                }
                i++;
            }
            int x = i - 1;
            for (int j = x; j >= rangeStart; j--) {
                theList.remove(j);
            }
            ObjectFactory factory = Context.getWmlObjectFactory();
            R run = factory.createR();
            Text t = factory.createText();
            t.setValue(value);
            run.getContent().add(t);
            theList.add(rangeStart, run);
        } catch (ClassCastException cce) {

        }
    }
}
