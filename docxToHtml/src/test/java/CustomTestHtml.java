import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * @author 孙茂林
 * @date 2022-06-10
 */
public class CustomTestHtml {

    private static final String sourceFile = "C:\\Users\\QM\\Desktop\\a.docx";
    private static final String targetFile = "C:\\Users\\QM\\Desktop\\a.html";

    @Test
    public void poiDocxNew() throws Exception {

        FileInputStream fileInputStream = new FileInputStream(sourceFile);

        XWPFDocument wordDocument = new XWPFDocument(fileInputStream);

        // 使用自定义图片管理。直接将文件上传至云服务器
        XHTMLOptions xhtmlOptions = XHTMLOptions.create().setImageManager(new CustomImageManager());

        XHTMLConverter.getInstance().convert(wordDocument, new FileOutputStream(targetFile), xhtmlOptions);

    }

}
