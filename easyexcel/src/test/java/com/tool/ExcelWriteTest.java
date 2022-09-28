package com.tool;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.util.ListUtils;
import com.tool.entity.DemoData;
import org.junit.Test;

import java.util.Date;
import java.util.List;

/**
 * @Description:
 */
public class ExcelWriteTest {
    
    String PATH = "D:\\IDEA\\IntelliJ IDEA 2021.3.3\\Workspace\\excel\\easyexcel\\";
    
    /* 通用数据生成 */
    private List<DemoData> data() {
        List<DemoData> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }
    
    /**
     * 1. 创建Excel对应的实体对象：{@link DemoData}
     * <p>
     * 2. 直接写即可
     */
    @Test
    public void testExcel() {
        /* 数据量不大的情况下可以使用（5000以内，具体也要看实际情况） */
        
        String fileName = PATH + "MyTable.xlsx";
        // 指定写Excel的文件名，用哪个Class去写，然后将数据写入到第一个名字为one的sheet，最后文件流会自动关闭
        EasyExcel.write(fileName, DemoData.class).sheet("one").doWrite(data());
    }
}