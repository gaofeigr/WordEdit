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
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Docx4jTest {

    static String TEMPLATE_DOCX_FILE_PATH = Docx4jTest.class.getClassLoader().getResource("template.docx").getFile();
    static String OUT_DOCX_FILE_PATH = Docx4jTest.class.getClassLoader().getResource("output.docx").getFile();
    static String RICH_TEXT_FILE_PATH = Docx4jTest.class.getClassLoader().getResource("richText.html").getFile();

    public static void main(String[] args) throws Exception {
//        // 创建一个空白word并插入文字
//        Docx4jUtil.testCreateDocx();
        // 拿到富文本内容(HTML代码)
        String richText = Files.readString(Path.of(RICH_TEXT_FILE_PATH), StandardCharsets.UTF_8);
//        // 使用html内容创建一个docx文件
//        Docx4jUtil.docx4jHtmlConvertToDocx(richText);
        // 替换富文本内容(书签方式替换)
        Docx4jUtil.replaceRichText(TEMPLATE_DOCX_FILE_PATH, richText, "replace3");
        // 替换普通文本(占位符方式替换), 用上面已经替换过富文本的文件再进行替换
        Map<String, String> valueMap = new HashMap<>();
        valueMap.put("replace1", "替换普通文本(占位符方式替换)");
        Docx4jUtil.replaceMapText(OUT_DOCX_FILE_PATH, valueMap);
        // 替换普通文本(书签方式替换), 用上面已经进行过替换的文件再进行替换
        Docx4jUtil.replaceBookMark(OUT_DOCX_FILE_PATH, "replace2", "书签值");
    }
}
