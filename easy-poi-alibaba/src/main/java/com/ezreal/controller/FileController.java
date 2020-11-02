package com.ezreal.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.fastjson.JSON;
import com.ezreal.model.FileModel;
import com.google.common.collect.Lists;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
public class FileController {



    @GetMapping("/api/file/download")
    public void download(HttpServletResponse response) {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        try {
            response.setContentType("application/vnd.ms-excel");
            response.setCharacterEncoding("utf-8");
            // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
            String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
            response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
            // 这里需要设置不关闭流
            List<FileModel> fileModelList = getFileModelList();
//            EasyExcel.write(response.getOutputStream(), FileModel.class).autoCloseStream(Boolean.FALSE).sheet("模板")
//                    .doWrite(fileModelList);

            ExcelWriterBuilder write = EasyExcel.write();
            write.file(response.getOutputStream());
            List<String> list = Lists.newArrayList("1", "2", "3");
            List<List<String>> head = Lists.newArrayList();
            head.add(list);
            write.head(head);
            write.autoCloseStream(Boolean.FALSE).sheet("模版").doWrite(fileModelList);

        } catch (Exception e) {
            // 重置response
            response.reset();
            response.setContentType("application/json");
            response.setCharacterEncoding("utf-8");
            Map<String, String> map = new HashMap<String, String>();
            map.put("status", "failure");
            map.put("message", "下载文件失败" + e.getMessage());
            try {
                response.getWriter().println(JSON.toJSONString(map));
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
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
