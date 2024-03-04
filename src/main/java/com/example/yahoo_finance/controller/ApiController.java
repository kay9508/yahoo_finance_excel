package com.example.yahoo_finance.controller;

import com.example.yahoo_finance.entity.FinanceInfo;
import com.example.yahoo_finance.service.FinanceService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@Slf4j
@RestController
@RequiredArgsConstructor
public class ApiController {

    private final FinanceService financeService;
    /*@GetMapping("/resultToExcel")
    public ResponseEntity<TodoResponseDto.TodoList> listTodo(@Valid TodoRequestDto.ListTodo listTodoDto,
                                                             @PageableDefault Pageable pageable) {
        TodoResponseDto.TodoList rtnDto = todoService.list(listTodoDto, pageable);
        return ResponseEntity
                .ok()
                .body(rtnDto);
    }*/

    @GetMapping("/downloadExcel")
    public void downloadExcel(HttpServletResponse response,  FinanceInfo financeInfo) throws IOException {

        financeService.test(financeInfo);
        /*
        // API 요청 처리 및 데이터 가져오기
        List<Data> data = ...;

        // Excel 파일 생성
        Workbook workbook = excelDownloadService.createExcel(data);

        // 파일 다운로드 설정
        response.setHeader("Content-Type", "application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=download.xlsx");

        // Excel 파일 쓰기
        workbook.write(response.getOutputStream());
        */
    }
}
