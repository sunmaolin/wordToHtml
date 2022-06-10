import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.junit.Before;
import org.junit.Test;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @author 孙茂林
 * @date 2022-06-10
 */
public class TestHtml {

    private static final String sourceFile = "C:\\Users\\QM\\Desktop\\a.doc";
    private static final String targetFile = "C:\\Users\\QM\\Desktop\\a.html";
    private static final String imageDir = "C:\\Users\\QM\\Desktop\\aImg";

    @Before
    public void mkdirImageDir() {
        File dir = new File(imageDir);
        if (!dir.exists()) {
            dir.mkdir();
        }
    }

    /**
     * .doc文档转html
     */
    @Test
    public void poiDoc() throws Exception {
        FileInputStream inputStream = new FileInputStream(sourceFile);

        HWPFDocument wordDocument = new HWPFDocument(inputStream);

        // HTML文档
        Document htmlDocument = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(htmlDocument);
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {
            @Override
            public String savePicture(byte[] content, PictureType pictureType, String name, float width, float height) {
                // 图片保存到桌面
                String path = imageDir + "\\" + name;
                try (FileOutputStream out = new FileOutputStream(path)) {
                    out.write(content);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                return path;
            }
        });
        // word转换为html输出到dom中
        wordToHtmlConverter.processDocument(wordDocument);

        // 获取dom转换源
        DOMSource domSource = new DOMSource(htmlDocument);
        // 获取流输出源
        StreamResult streamResult = new StreamResult(new File(targetFile));

        // java中从源到结果的转换
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        // 设置属性
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.METHOD, "html");
        // 转换
        transformer.transform(domSource, streamResult);

//        org.jsoup.nodes.Document parse = Jsoup.parse(new File(targetFile), "UTF-8");
//        Elements html = parse.getElementsByTag("html");
//        System.out.println(html.html());

    }

}
