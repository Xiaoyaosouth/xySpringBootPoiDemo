package com.fe.mysbpoidemo.controller;

import com.fe.mysbpoidemo.service.DynamicExcelDataService;
import org.springframework.web.bind.annotation.*;

import java.util.*;

/**
 * 控制器 API调用
 * @author rk
 */
@RestController
@RequestMapping("/xy")
public class DynamicExcelReaderController {

    /**
     * 解析Excel文件
     * @param fileUrl 文件网络地址
     * @return
     */
    @RequestMapping(value = "/dynamicExcelData", method = RequestMethod.GET, params = "fileUrl")
    public List<Object> getMyDynamicExcelData(String fileUrl) {
        System.out.println("-------获取动态生成的表格数据-------");
        DynamicExcelDataService myService = new DynamicExcelDataService();
        // 使用新方法，处理中间空列空行无法取值的问题 2020-05-08
        List<Object> data = myService.getDynamicModelListV2(fileUrl);
        return data;
    }
}
