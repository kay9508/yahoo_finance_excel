package com.example.yahoo_finance.controller;

import com.example.yahoo_finance.entity.FinanceInfo;
import com.example.yahoo_finance.service.FinanceService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.net.URLEncoder;
import java.rmi.ServerException;

@Slf4j
@RestController
@RequiredArgsConstructor
public class ApiController {

    private final FinanceService financeService;
    @GetMapping("/downloadExcel")
    public void downloadExcel(HttpServletResponse response,  FinanceInfo financeInfo) throws IOException {

        String fileName = "";
        if (financeInfo.getTicker() == null) {
            throw new IOException("Primary value is not Input");
        } else {
            if (financeInfo.getShotTag() != null && !financeInfo.getShotTag().isEmpty()) {
                fileName += "(" + financeInfo.getShotTag() + ") ";
            }
            fileName += financeInfo.getTicker();
        }

        XSSFWorkbook workbook = financeService.excelDownload(financeInfo);
        if (workbook == null) {
            throw new ServerException("Excel Download Error");
        }
        response.setContentType("application/vnd.ms-excel");
        String checkName = URLEncoder.encode(fileName, "UTF-8");
        //response.setHeader("Content-Disposition", "attachment;filename="+ URLEncoder.encode(fileName, "UTF-8")+".xlsx");
        //파일명은 URLEncoder로 감싸주는게 좋다! <- 감싸면 특수문자 깨져서 아래처럼 일단 사용
        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + ".xlsx\"");

        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
