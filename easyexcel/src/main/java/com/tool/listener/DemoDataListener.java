package com.tool.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.fastjson2.JSON;
import com.tool.dao.DemoDao;
import com.tool.entity.DemoData;

import java.util.List;

/**
 * @Description:
 */
// 注意：DemoDataListener不能被Spring管理，每次读取Excel都要new，然后里面用到Spring可以构造方法传进去
public class DemoDataListener implements ReadListener<DemoData> {
    
    /**
     * 每隔5条存储到数据库，实际使用中可以100条，然后清理list，方便内存回收
     */
    private static final int BATCH_COUNT = 100;
    
    /**
     * 缓存的数据
     */
    private List<DemoData> cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);
    
    /**
     * 假设这个是一个DAO，当然有业务逻辑这个也可以是一个service。当然如果不用存储这个对象没用。
     */
    private DemoDao demoDao;
    
    public DemoDataListener() {
        // 这里是demo，所以随便new一个。实际使用中如果用到了Spring，请使用下面的有参构造函数
        demoDao = new DemoDao();
    }
    
    /**
     * 如果使用了Spring，请使用这个有参构造方法。每次创建Listener的时候需要把Spring管理的类传进来
     *
     * @param demoDao
     */
    public DemoDataListener(DemoDao demoDao) {
        this.demoDao = demoDao;
    }
    
    /**
     * 每一条数据解析都会调用，即读取数据时都会执行该方法
     *
     * @param data    one row value. Is is same as {@link AnalysisContext#readRowHolder()}
     * @param context
     */
    @Override
    public void invoke(DemoData data, AnalysisContext context) {
        System.out.println("解析到一条数据：" + JSON.toJSONString(data));
        cachedDataList.add(data);
        // 如果缓存的数据达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (cachedDataList.size() >= BATCH_COUNT) {
            saveData();
            // 存储完成清理list
            cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);
        }
    }
    
    /**
     * 所有数据解析完成都会调用
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
        saveData();
        System.out.println("所有数据解析完成！");
    }
    
    /**
     * 加上存储数据库
     */
    private void saveData() {
        System.out.println(cachedDataList.size() + "条数据，开始存储数据库！");
        demoDao.save(cachedDataList);
        System.out.println("存储数据库成功！");
    }
}
