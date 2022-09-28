package com.tool.dao;

import com.tool.entity.DemoData;

import java.util.List;

/**
 * 假设这个是DAO存储，需要让这个类给Spring管理；如果不用需要存储，则不需要这个类。
 **/
public class DemoDao {
    public void save(List<DemoData> list) {
        // 如果是MyBatis，尽量别直接调用多次insert，自己写一个mapper里面新增一个方法batchInsert，将所有数据一次性插入
    }
}