package org.example;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.finders.RangeFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.*;

import java.io.File;
import java.util.List;
import java.util.Map;

public class Docx4jUtil {

    /**
     * 根据占位符替换docx文件中的文字
     * @param valueMap 占位符:替换的值     例: ${test1}:替换值
     * @param templateFilePath 模板文件路径
     * @throws Exception
     */
    public static void replaceMapText(String templateFilePath, Map<String, String> valueMap) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(templateFilePath));
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
        mainDocumentPart.variableReplace(valueMap);
        wordMLPackage.save(new File(Docx4jTest.OUT_DOCX_FILE_PATH));
    }

    /**
     * 将富文本替换到docx文件中
     * @param templateFilePath 模板文件路径
     * @param richText 富文本内容
     * @param bookMark 书签名称
     * @return
     */
    public static void replaceRichText(String templateFilePath, String richText, String bookMark) throws Docx4JException {
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
        replaceTemplate(templateFilePath, sbf.toString(), bookMark);
    }

    private static void replaceTemplate(String templateFilePath, String richContent, String bookMark) throws Docx4JException {
        WordprocessingMLPackage aPackage;
        try {
            aPackage = WordprocessingMLPackage.load(new File(templateFilePath));
            MainDocumentPart mainDocumentPart = aPackage.getMainDocumentPart();
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
                if (bm.getName().equals(bookMark)) {
                    // 获取该书签的父级段落
                    P p = (P) (bm.getParent());
                    paragraphs.addAll(paragraphs.indexOf(p), importer.convert(richContent, null));
                    paragraphs.remove(p);
                }
            }
            aPackage.save(new File(Docx4jTest.OUT_DOCX_FILE_PATH));
        } catch (Exception e) {
            throw e;
        }
    }

    /**
     * 替换书签内容
     * @param templateFilePath 模板文件路径
     * @param bookMark 书签名称
     * @param value 替换值
     */
    public static void replaceBookMark(String templateFilePath, String bookMark, String value) throws Exception {
        WordprocessingMLPackage aPackage;
        try {
            aPackage = WordprocessingMLPackage.load(new File(templateFilePath));
            MainDocumentPart mainDocumentPart = aPackage.getMainDocumentPart();
            // 提取正文中所有段落
            List<Object> paragraphs = mainDocumentPart.getContent();
            // 提取书签并创建书签的游标
            RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");
            new TraversalUtil(paragraphs, rt);
            // 遍历书签
            for (CTBookmark bm : rt.getStarts()) {
                // 这儿可以对单个书签进行操作，也可以用一个map对所有的书签进行处理
                if (bm.getName().equals(bookMark)) {
                    replaceText(bm, value);
                }
            }
            aPackage.save(new File(Docx4jTest.OUT_DOCX_FILE_PATH));
        } catch (Exception e) {
            throw e;
        }
    }

    /**
     * 替换书签内容
     * @param bm
     * @param object
     * @throws Exception
     */
    private static void replaceText(CTBookmark bm, Object object) throws Exception {
        if (object == null) {
            return;
        }
        if (bm.getName() == null){
            return;
        }
        String value = object.toString();
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
    }

    /**
     * 测试生成docx文件并写入文字
     * @throws Docx4JException
     */
    public static void testCreateDocx() throws Docx4JException {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(Docx4jTest.OUT_DOCX_FILE_PATH));
        File file = new File(Docx4jTest.OUT_DOCX_FILE_PATH);
        wordMLPackage.getMainDocumentPart().addParagraphOfText("你好!");
        wordMLPackage.save(file);
    }

    /**
     * 将html转换为docx文件
     * @param html 富文本html
     * @throws Exception
     */
    public static void docx4jHtmlConvertToDocx(String html) throws Exception {
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
        sbf.append(html);
        sbf.append("</body></html>");
        WordprocessingMLPackage pack = WordprocessingMLPackage.createPackage();
        NumberingDefinitionsPart numDefPart = new NumberingDefinitionsPart();
        pack.getMainDocumentPart().addTargetPart(numDefPart);
        numDefPart.unmarshalDefaultNumbering();
        // Convert XHTML + CSS to WordML content.
        // XHTML must be well formed XML.
        XHTMLImporterImpl importer = new XHTMLImporterImpl(pack);
        // Convert the well formed XHTML contained in file to a list of WML objects.
        List<Object> list = importer.convert(sbf.toString(), null);
        MainDocumentPart mainPart = pack.getMainDocumentPart();
        // Add all WML objects to MainDocumentPart.
        mainPart.getContent().addAll(list);
        // set compatibilityMode to 15
        // to avoid Word 365/2016 saying "Compatibility Mode"
        DocumentSettingsPart settingsPart = mainPart.getDocumentSettingsPart(true);
        CTCompat compat = Context.getWmlObjectFactory().createCTCompat();
        compat.setCompatSetting("compatibilityMode", "http://schemas.microsoft.com/office/word", "15");
        settingsPart.getContents().setCompat(compat);
        // save to a file
        pack.save(new File(Docx4jTest.OUT_DOCX_FILE_PATH));
    }
}
