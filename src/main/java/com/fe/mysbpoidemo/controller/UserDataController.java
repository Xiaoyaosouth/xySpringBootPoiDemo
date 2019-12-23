package com.fe.mysbpoidemo.controller;

import com.fe.mysbpoidemo.*;
import com.fe.mysbpoidemo.model.*;
import com.fe.mysbpoidemo.service.*;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

/**
 * @author rk
 */
@RestController
@RequestMapping("/xy")
    public class UserDataController {

    @RequestMapping(value = "/myUser", method = RequestMethod.GET)
    public UserXy getUserXy() {
        System.out.println("-------DIY-------");
        UserXy user = new UserXy();
        user.setName("逍遥");
        user.setAge("22");
        user.setPhone("13750002413");
        return user;
    }

    @RequestMapping(value = "/excelData", method = RequestMethod.GET, params = "fileUrl")
    public List<UserXy> getUserXyListData(String fileUrl) {
        System.out.println("-------获取表格数据-------");
        UserDataService userDataService = new UserDataService();
        List<UserXy> data = userDataService.getModelList(fileUrl);
        return data;
    }

    @RequestMapping(value = "/excelData2", method = RequestMethod.GET, params = "fileUrl")
    public ResponseResult<List<UserXy>> getUserXyListData2(String fileUrl) {
        try {
            System.out.println("-------获取表格数据2-------");
            UserDataService userDataService = new UserDataService();
            List<UserXy> data = userDataService.getModelList(fileUrl);
            return Response.makeOKRsp(data);
        } catch (Exception e) {
            e.printStackTrace();
            return Response.makeErrRsp("查询用户信息异常");
        }

    }

}
