package com.tool;

import com.alibaba.excel.EasyExcel;
import com.tool.entity.DemoData;
import com.tool.listener.DemoDataListener;
import org.junit.Test;

/**
 * @Description:
 */
public class ExcelReadTest {
    
    String PATH = "D:\\IDEA\\IntelliJ IDEA 2021.3.3\\Workspace\\excel\\easyexcel\\";
    
    /**
     * <p>1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>2. 由于默认一行行的读取excel，所以需要创建excel一行一行的回调监听器，参照{@link DemoDataListener}
     * <p>3. 直接读即可
     */
    @Test
    public void testExcel() {
        String fileName = PATH + "MyTable.xlsx";
        // 需要指定Excel的文件名，用哪个class去读，然后读取到第一个sheet，最后文件流会自动关闭
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
    }
}
