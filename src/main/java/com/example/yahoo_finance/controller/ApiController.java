package com.example.yahoo_finance.controller;

import com.example.yahoo_finance.entity.FinanceInfo;
import com.example.yahoo_finance.service.FinanceService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

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

    @PostMapping("/uploadExcel")
    public void uploadExcel(@RequestBody MultipartFile file) throws IOException {
        XSSFWorkbook xlsx = new XSSFWorkbook(file.getInputStream());
        XSSFSheet sheet1 = xlsx.getSheetAt(0);
        XSSFSheet sheet2 = xlsx.getSheetAt(1);

        CellStyle row1cell1Style = sheet2.getRow(0).getCell(0).getCellStyle();
//        CellStyle row1cell2Style = sheet2.getRow(0).getCell(1).getCellStyle();
        CellStyle row1cell3Style = sheet2.getRow(0).getCell(2).getCellStyle();
        CellStyle row1cell4Style = sheet2.getRow(0).getCell(3).getCellStyle();
        CellStyle row1cell5Style = sheet2.getRow(0).getCell(4).getCellStyle();
        CellStyle row1cell6Style = sheet2.getRow(0).getCell(5).getCellStyle();
        CellStyle row1cell7Style = sheet2.getRow(0).getCell(6).getCellStyle();
        CellStyle row1cell8Style = sheet2.getRow(0).getCell(7).getCellStyle();

        CellStyle row2cell1Style = sheet2.getRow(1).getCell(0).getCellStyle();
        CellStyle row2cell2Style = sheet2.getRow(1).getCell(1).getCellStyle();
        CellStyle row2cell3Style = sheet2.getRow(1).getCell(2).getCellStyle();
        CellStyle row2cell4Style = sheet2.getRow(1).getCell(3).getCellStyle();
        CellStyle row2cell5Style = sheet2.getRow(1).getCell(4).getCellStyle();
        CellStyle row2cell6Style = sheet2.getRow(1).getCell(5).getCellStyle();
        CellStyle row2cell7Style = sheet2.getRow(1).getCell(6).getCellStyle();
        CellStyle row2cell8Style = sheet2.getRow(1).getCell(7).getCellStyle();

        CellStyle row3cell1Style = sheet2.getRow(2).getCell(0).getCellStyle();
        CellStyle row3cell2Style = sheet2.getRow(2).getCell(1).getCellStyle();
        CellStyle row3cell3Style = sheet2.getRow(2).getCell(2).getCellStyle();
        CellStyle row3cell4Style = sheet2.getRow(2).getCell(3).getCellStyle();
        CellStyle row3cell5Style = sheet2.getRow(2).getCell(4).getCellStyle();
        CellStyle row3cell6Style = sheet2.getRow(2).getCell(5).getCellStyle();
        CellStyle row3cell7Style = sheet2.getRow(2).getCell(6).getCellStyle();
        CellStyle row3cell8Style = sheet2.getRow(2).getCell(7).getCellStyle();

        XSSFWorkbook xls = new XSSFWorkbook(); // .xlsx
        CellStyle copyCellStyle1 = xls.createCellStyle();

        copyCellStyle1.cloneStyleFrom(row1cell1Style);

        System.out.println("test");

    }
}
