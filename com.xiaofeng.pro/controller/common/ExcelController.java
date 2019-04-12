package com.xiaofeng.pro.controller.common;

import com.xiaofeng.pro.common.utils.ExportExcelUtil;
import com.xiaofeng.pro.entity.Organization;
import com.xiaofeng.pro.service.OrganizationService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

/**
 * excel通用
 * 2019.3.14
 */
@RestController
@RequestMapping("/edu/excel")
@Slf4j
public class ExcelController {

    @Resource
    private OrganizationService organizationService;

    @GetMapping("/testexcel")
    public void testExcel(HttpServletResponse response) {
        List<Organization> organizations = organizationService.getAll();
        try {
            new ExportExcelUtil("just test", Organization.class).setExcelData(organizations).write(response, "xiaofeng.xls").dispose();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
