import fr.opensagres.poi.xwpf.converter.core.ImageManager;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * 自定义word转html图片处理
 * 源码实现方式不符合预期。继承重写主要方法
 * @author 孙茂林
 * @date 2022-06-10
 */
public class CustomImageManager extends ImageManager {

    /**
     * 根据插件自动生成的地址与上传到服务器的地址做映射缓存
     */
    Map<String, String> imageUrlCache = new HashMap<>();

    public CustomImageManager() {
        super(null, null);
    }

    @Override
    public String resolve(String uri) {
        return imageUrlCache.get(uri);
    }

    @Override
    public void extract(String imagePath, byte[] imageData) throws IOException {
        /*imagePath example: word/media/image1.png word/media/image2.png */
        // 1.保存图片到服务器返回访问链接
        String serverImageUrl = "https://p3-juejin.byteimg.com/tos-cn-i-k3u1fbpfcp/3535ab1b9d4a45ceb9369e56c697e07b~tplv-k3u1fbpfcp-watermark.image";
        // 2.做映射缓存
        imageUrlCache.put(imagePath, serverImageUrl);
    }
}
