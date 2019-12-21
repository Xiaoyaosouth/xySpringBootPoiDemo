package com.fe.mysbpoidemo.controller;

import com.fe.mysbpoidemo.model.ExcelCommon;
import com.fe.mysbpoidemo.service.CommonExcelReaderService;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/xy")
public class CommonExcelReaderController {

    @RequestMapping(value = "/commonExcelData", method = RequestMethod.GET, params = "fileUrl")
    public ExcelCommon getCommonExcelData(String fileUrl) {
        System.out.println("-------获取通用表格数据-------");
        CommonExcelReaderService service = new CommonExcelReaderService();
        ExcelCommon excelCommon = service.getCommonExcelData(fileUrl);
        return excelCommon;
    }

}
