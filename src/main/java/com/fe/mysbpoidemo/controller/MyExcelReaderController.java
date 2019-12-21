package com.fe.mysbpoidemo.controller;

import com.fe.mysbpoidemo.model.ExcelCommon;
import com.fe.mysbpoidemo.service.ExcelDataService;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.util.ArrayList;
import java.util.List;

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

    @RequestMapping(value = "/testCommonExcelData", method = RequestMethod.GET)
    public ExcelCommon getMyCommonExcelData() {
        System.out.println("-------获取通用表格数据-------");

        List<String> stringList = new ArrayList<>();
        stringList.add("name");
        stringList.add("phone");

        List<String> nameValue = new ArrayList<>();
        nameValue.add("逍遥科技");
        nameValue.add("大疆科技");

        List<String> phoneValue = new ArrayList<>();
        phoneValue.add("13750002413");
        phoneValue.add("13751234321");

        List<List<String>> cellData = new ArrayList<>();
        cellData.add(nameValue);
        cellData.add(phoneValue);

        ExcelCommon data = new ExcelCommon();
        data.setProperty(stringList);
        data.setCellData(cellData);


        return data;
    }

}
