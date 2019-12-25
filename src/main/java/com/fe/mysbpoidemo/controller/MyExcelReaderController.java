package com.fe.mysbpoidemo.controller;

import com.fe.mysbpoidemo.service.ExcelDataService;
import org.springframework.web.bind.annotation.*;

import java.util.*;

/**
 * 控制器
 * @author rk
 */
@RestController
@RequestMapping("/xy")
public class MyExcelReaderController {

    @RequestMapping(value = "/dynamicExcelData", method = RequestMethod.GET, params = "fileUrl")
    public List<Object> getMyDynamicExcelData(String fileUrl) {
        System.out.println("-------获取动态生成的表格数据-------");
        ExcelDataService excelDataService = new ExcelDataService();
        List<Object> data = excelDataService.getDynamicModelList(fileUrl);
        return data;
    }

}
