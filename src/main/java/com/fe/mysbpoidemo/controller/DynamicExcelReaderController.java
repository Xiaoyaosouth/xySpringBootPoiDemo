package com.fe.mysbpoidemo.controller;

import com.fe.mysbpoidemo.service.DynamicExcelDataService;
import org.springframework.web.bind.annotation.*;

import java.util.*;

/**
 * 控制器
 * @author rk
 */
@RestController
@RequestMapping("/xy")
public class DynamicExcelReaderController {

    /**
     *
     * @param fileUrl 文件网络地址
     * @return
     */
    @RequestMapping(value = "/dynamicExcelData", method = RequestMethod.GET, params = "fileUrl")
    public List<Object> getMyDynamicExcelData(String fileUrl) {
        System.out.println("-------获取动态生成的表格数据-------");
        DynamicExcelDataService myService = new DynamicExcelDataService();
        List<Object> data = myService.getDynamicModelList(fileUrl);
        return data;
    }

}
