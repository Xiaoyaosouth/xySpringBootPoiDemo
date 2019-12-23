package com.fe.mysbpoidemo.controller;

import com.fe.mysbpoidemo.model.*;
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

    @RequestMapping(value = "/commonExcelDataV2", method = RequestMethod.GET, params = "fileUrl")
    public ExcelCommonV2 getCommonExcelDataV2(String fileUrl) {
        System.out.println("-------获取通用表格数据V2-------");
        CommonExcelReaderService service = new CommonExcelReaderService();
        ExcelCommonV2 data = service.getCommonExcelDataV2(fileUrl);
        return data;
    }

}
