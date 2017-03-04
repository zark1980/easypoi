import com.telecomjs.utils.PoiExcelWriter;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by zark on 17/3/4.
 */
public class TestPoiExcelWriter {

    @Test()
    public void testExcel() throws FileNotFoundException {
        FileOutputStream os = new FileOutputStream("/Users/zark/Downloads/doc2.xlsx");
        PoiExcelWriter writer = new PoiExcelWriter(os,"xlsx");
        writer.open();
        String[] titles = {"col1","col2","col3","col4","col5","col6","col7","col8","col9","col10"};
        writer.writeTitle(titles);
        for (int i=0;i<10;i++){
            String[] strs = {"1","2","3","4","5","6","7","8","9","10"};
            writer.writeRow(strs);
        }
        writer.close();
    }

    @Test()
    public void testExcelAll() throws FileNotFoundException {
        FileOutputStream os = new FileOutputStream("/Users/zark/Downloads/doc3.xlsx");
        PoiExcelWriter writer = new PoiExcelWriter(os,"xlsx");
        writer.open();
        String[] titles = {"col1","col2","col3","col4","col5","col6","col7","col8","col9","col10"};
        writer.writeTitle(titles);
        List list = new ArrayList();
        for (int i=0;i<20;i++){
            String[] strs = {"1","2","3","4","5","6","7","8","9","10"};
            list.add(strs);
        }
        writer.writeAll(list);
        writer.close();
    }


}
