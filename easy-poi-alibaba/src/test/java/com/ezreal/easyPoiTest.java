package com.ezreal;

import com.alibaba.excel.EasyExcel;
import com.ezreal.model.FileModel;
import com.ezreal.util.TestFileUtil;
import com.google.common.collect.Lists;
import org.junit.jupiter.api.Test;

import java.util.List;

public class easyPoiTest {

    /**
     * 最简单的写
     */
    @Test
    public void simpleWrite() {
        String fileName = TestFileUtil.getPath() + "write" + System.currentTimeMillis() + ".xlsx";
        System.out.println(fileName);
        // 这里 需要指定写用哪个class去读，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
        EasyExcel.write(fileName, FileModel.class).sheet("模板").doWrite(getFileModelList());
    }

    private List<FileModel> getFileModelList(){
        List<FileModel> fileModelList = Lists.newArrayList();
        for (int i = 0; i < 10; i++) {
            FileModel fileModel = new FileModel();
            fileModel.setFileName("fileName_"+i);
            fileModel.setFileUrl("fileUrl_"+i);
            fileModelList.add(fileModel);
        }
        return fileModelList;
    }
}
