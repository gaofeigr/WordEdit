package org.example;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;

public class Docx4jTest {

    static String TEMPLATE_DOCX_FILE_PATH = Docx4jTest.class.getClassLoader().getResource("template2.docx").getFile();
    static String OUT_DOCX_FILE_PATH = Docx4jTest.class.getClassLoader().getResource("output.docx").getFile();
    static String RICH_TEXT_FILE_PATH = Docx4jTest.class.getClassLoader().getResource("richText.html").getFile();

    public static void main(String[] args) throws Exception {
//        // 创建一个空白word并插入文字·
//        Docx4jUtil.testCreateDocx();
        // 拿到富文本内容(HTML代码)
        String richText = getHtmlFileString(RICH_TEXT_FILE_PATH);
////        // 使用html内容创建一个docx文件
//        Docx4jUtil.docx4jHtmlConvertToDocx(richText);
        // 替换富文本内容(书签方式替换)
        Docx4jUtil.replaceRichText(TEMPLATE_DOCX_FILE_PATH, richText, "replace3");
        // 替换普通文本(占位符方式替换), 用上面已经替换过富文本的文件再进行替换
//        Map<String, String> valueMap = new HashMap<>();
//        valueMap.put("replace1", "替换普通文本(占位符方式替换)");
//        Docx4jUtil.replaceMapText(OUT_DOCX_FILE_PATH, valueMap);
//         替换普通文本(书签方式替换), 用上面已经进行过替换的文件再进行替换
//        Docx4jUtil.replaceBookMark(OUT_DOCX_FILE_PATH, "replace2", "书签值");
    }

    private static String getHtmlFileString(String htmlFilePath) {
        StringBuffer result = new StringBuffer();
        try (Scanner sc = new Scanner(new FileReader(htmlFilePath))) {
            while (sc.hasNextLine()) {  //按行读取字符串
                result.append(sc.nextLine());
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return result.toString();
    }
}
