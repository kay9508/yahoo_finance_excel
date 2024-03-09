package com.example.yahoo_finance.service;

import com.example.yahoo_finance.entity.ApiResultDTO;
import com.example.yahoo_finance.entity.FinanceInfo;
import com.example.yahoo_finance.entity.PeriodDTO;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.web.reactive.function.client.WebClient;

import java.sql.Timestamp;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.LinkedHashMap;
import java.util.List;

@Slf4j
@Service
@RequiredArgsConstructor
public class FinanceService {

    @Value("${yahoo.finance.url}")
    private String yahooFinanceUrl;

    @Value("${yahoo.finance.api.key}")
    private String yahooFinanceApiKey;

    private final ObjectMapper objectMapper;

    private final WebClient webClient;


    public XSSFWorkbook excelDownload(FinanceInfo financeInfo) {
        String companyFullName = getCompanyFullName(financeInfo);
        PeriodDTO periodSetting = getPeriods(financeInfo);
        String dayDataUrl = getRequestedUrl(financeInfo.getTicker(), periodSetting, "1d");
        String monthDataUrl = getRequestedUrl(financeInfo.getTicker(), periodSetting, "1mo");
        String nbiUrl = getRequestedUrl("^NBI", periodSetting, "1d");

        String nbiResponse = "";
        try {
            nbiResponse = webClient.get()
                    .uri(nbiUrl)
                    //.header("Authorization", "Bearer " + yahooFinanceApiKey)
                    .retrieve()
                    .bodyToMono(String.class)
                    .block();
        } catch (RuntimeException e) {
            // 예외 처리
            log.error("파이낸셜 api로 정보를 가져오던 중 에러가 발생했습니다.");
            e.printStackTrace(); // 예외 정보 출력
            return null;
        }
        ApiResultDTO nbiResultDTO = null;
        try {
            JSONObject jsonObject = objectMapper.readValue(nbiResponse, JSONObject.class);
            LinkedHashMap chartMap = (LinkedHashMap) jsonObject.get("chart");
            List resultList = (List) chartMap.get("result");
            LinkedHashMap resultMap = (LinkedHashMap) resultList.get(0);

            nbiResultDTO = objectMapper.convertValue(resultMap, ApiResultDTO.class);
            log.info("resultDTO : {}", nbiResultDTO);
        } catch (Exception e) {
            log.error("DTO변환 도중 에러가 발생했습니다.");
            e.printStackTrace();
        }

        LinkedHashMap<LocalDate, Double> nbiMap = new LinkedHashMap();
        for (int i = 0; i < nbiResultDTO.getTimestamp().size(); i++) {
            LocalDate tradeDate = Instant.ofEpochSecond(nbiResultDTO.getTimestamp().get(i)).atZone(ZoneId.systemDefault()).toLocalDate();
            Double adjClose = nbiResultDTO.getIndicators().getAdjclose().get(0).getAdjclose().get(i);
            nbiMap.put(tradeDate, Double.valueOf(String.format("%.2f", adjClose)));
        }

        String dayDataResponse = "";
        try {

            dayDataResponse = webClient.get()
                    .uri(dayDataUrl)
                    //.header("Authorization", "Bearer " + yahooFinanceApiKey)
                    .retrieve()
                    .bodyToMono(String.class)
                    .block();
        } catch (RuntimeException e) {
            // 예외 처리
            log.error("파이낸셜 api로 정보를 가져오던 중 에러가 발생했습니다.");
            e.printStackTrace(); // 예외 정보 출력
            return null;
        }

        ApiResultDTO dayDataResultDTO = null;
        try {
            JSONObject jsonObject = objectMapper.readValue(dayDataResponse, JSONObject.class);
            LinkedHashMap chartMap = (LinkedHashMap) jsonObject.get("chart");
            List resultList = (List) chartMap.get("result");
            LinkedHashMap resultMap = (LinkedHashMap) resultList.get(0);

            dayDataResultDTO = objectMapper.convertValue(resultMap, ApiResultDTO.class);
            log.info("resultDTO : {}", dayDataResultDTO);
        } catch (Exception e) {
            log.error("DTO변환 도중 에러가 발생했습니다.");
            e.printStackTrace();
        }

        /*LocalDate firstTradeDate = Instant.ofEpochSecond(dayDataResultDTO.getMeta().getFirstTradeDate()).atZone(ZoneId.systemDefault()).toLocalDate();
        if (!firstTradeDate.equals(financeInfo.getStartDate())) {
            log.warn("(!주의!)첫 거래일이 입력된 일자와 다릅니다.");
        }*/

        XSSFWorkbook makedExcel = excelCreateAndSetDayData(dayDataResultDTO, financeInfo.getShotTag(), financeInfo.getKoreaName(), nbiMap);
        XSSFWorkbook addedExcel = setMonthData(makedExcel, monthDataUrl, companyFullName, financeInfo.getTicker());
        log.info("끝");
        return addedExcel;
    }

    private XSSFWorkbook excelCreateAndSetDayData(ApiResultDTO resultDTO, String shotTag, String koreaName, LinkedHashMap<LocalDate, Double> nbiMap) {
        shotTag = shotTag == null ? "" : shotTag;
        koreaName = koreaName == null ? "" : koreaName;

        List<Double> closeList = resultDTO.getIndicators().getQuote().get(0).getClose();
        List<Integer> volumeList = resultDTO.getIndicators().getQuote().get(0).getVolume();
        List<Double> openList = resultDTO.getIndicators().getQuote().get(0).getOpen();
        List<Double> lowList = resultDTO.getIndicators().getQuote().get(0).getLow();
        List<Double> highList = resultDTO.getIndicators().getQuote().get(0).getHigh();

        // 엑셀 파일 생성
        XSSFWorkbook xls = new XSSFWorkbook(); // .xlsx

        Font defaultFont = xls.createFont();
        defaultFont.setFontName("Arial");
        defaultFont.setFontHeightInPoints((short) 15);

        // 기본 스타일 (데이터)
        CellStyle defaultFontStyle = xls.createCellStyle();
        defaultFontStyle.setFont(defaultFont);

        // A-1에 사용될 스타일
        CellStyle row1Cell1Style = xls.createCellStyle();
        Font row1Cell1Font = xls.createFont();
        row1Cell1Font.setFontName("Arial");
        row1Cell1Font.setBold(true);
        row1Cell1Font.setFontHeightInPoints((short) 17);
        row1Cell1Style.setFont(row1Cell1Font);

        // H-1에 사용될 스타일
        CellStyle row1Cell8Style = xls.createCellStyle();
        Font row1Cell8Font = xls.createFont();
        row1Cell8Font.setFontName("Arial");
        row1Cell8Font.setBold(true);
        row1Cell8Font.setFontHeightInPoints((short) 14);
        row1Cell8Style.setFont(row1Cell8Font);

        // A-2에 사용될 스타일
        CellStyle row2Cell1Style = xls.createCellStyle();
        Font row2Cell1Font = xls.createFont();
        row2Cell1Font.setFontName("Arial");
        row2Cell1Font.setBold(true);
        row2Cell1Font.setFontHeightInPoints((short) 15);
        row2Cell1Style.setFont(row2Cell1Font);
        row2Cell1Style.setBorderTop(BorderStyle.THIN);
        row2Cell1Style.setBorderLeft(BorderStyle.THIN);
        row2Cell1Style.setBorderRight(BorderStyle.DOUBLE);
        row2Cell1Style.setBorderBottom(BorderStyle.DOUBLE);

        // row2에 사용될 스타일
        CellStyle row2Cell2to8Style = xls.createCellStyle();
        Font row2Cell2to8Font = xls.createFont();
        row2Cell2to8Font.setFontName("Arial");
        row2Cell2to8Font.setBold(true);
        row2Cell2to8Font.setFontHeightInPoints((short) 15);
        row2Cell2to8Style.setFont(row2Cell2to8Font);
        row2Cell2to8Style.setBorderTop(BorderStyle.THIN);
        row2Cell2to8Style.setBorderLeft(BorderStyle.THIN);
        row2Cell2to8Style.setBorderRight(BorderStyle.THIN);
        row2Cell2to8Style.setBorderBottom(BorderStyle.DOUBLE);

        // VOLUME에 사용될 스타일
        CellStyle cashStyle = xls.createCellStyle();
        cashStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0"));
        cashStyle.setFont(defaultFont);
        cashStyle.setBorderRight(BorderStyle.THIN);
        cashStyle.setBorderBottom(BorderStyle.THIN);

        // CHANGE에 사용될 스타일
        CellStyle percentStyle = xls.createCellStyle();
        percentStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00%"));
        percentStyle.setFont(defaultFont);
        percentStyle.setBorderRight(BorderStyle.THIN);
        percentStyle.setBorderBottom(BorderStyle.THIN);

        // NBI에 사용될 스타일
        CellStyle cashDotTwoStyle = xls.createCellStyle();
        cashDotTwoStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
        cashDotTwoStyle.setFont(defaultFont);
        cashDotTwoStyle.setBorderLeft(BorderStyle.THIN);
        cashDotTwoStyle.setBorderBottom(BorderStyle.THIN);

        // 나머지 Data에 사용될 스타일
        CellStyle dotTwoStyle = xls.createCellStyle();
        dotTwoStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        dotTwoStyle.setFont(defaultFont);
        dotTwoStyle.setBorderRight(BorderStyle.THIN);
        dotTwoStyle.setBorderBottom(BorderStyle.THIN);

        // 기본 스타일
        CellStyle defaultStyle = xls.createCellStyle();
        defaultStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        defaultStyle.setFont(defaultFont);
        defaultStyle.setBorderTop(BorderStyle.THIN);
        defaultStyle.setBorderLeft(BorderStyle.THIN);
        defaultStyle.setBorderRight(BorderStyle.THIN);
        defaultStyle.setBorderBottom(BorderStyle.THIN);

        // 일자에 사용될 스타일
        CellStyle dayStyle = xls.createCellStyle();
        dayStyle.setFont(defaultFont);
        dayStyle.setBorderLeft(BorderStyle.DOUBLE);
        dayStyle.setBorderRight(BorderStyle.THIN);
        dayStyle.setBorderBottom(BorderStyle.THIN);

        // 내림차순으로 정렬 (이거 하면 데이터 다꼬임)
        // Collections.sort(resultDTO.getTimestamp(), Collections.reverseOrder());

        LinkedHashMap<String, XSSFSheet> sheetMap = new LinkedHashMap<>();

        List<Integer> timestampList = resultDTO.getTimestamp();
        Integer totalCount = resultDTO.getTimestamp().size();
        for (Integer i = totalCount - 1; i >= 0; i--) {
            LocalDate tradeDate = Instant.ofEpochSecond(timestampList.get(i)).atZone(ZoneId.systemDefault()).toLocalDate();
            String sheetName = String.valueOf(tradeDate.getMonth()).substring(0,3) + " " + String.valueOf(tradeDate.getYear()).substring(2,4);

            //  시트생성
            if (sheetMap.get(sheetName) == null) {
                XSSFSheet sheet = xls.createSheet(sheetName);
                sheet.setMargin(PageMargin.TOP, Double.valueOf("0.6")); // 여백 위쪽1.5
                sheet.setMargin(PageMargin.LEFT, Double.valueOf("0.5")); // 여백 왼쪽1.3
                sheet.setMargin(PageMargin.RIGHT, Double.valueOf("0.5")); // 여백 오른쪽1.3
                sheet.setMargin(PageMargin.BOTTOM, Double.valueOf("0.3")); // 여백 아래쪽0.8
                sheet.setHorizontallyCenter(true); // 페이지 가운데 맞춤
                // sheet.setZoom(100);

                // HSSFPrintSetup 또는 XSSFPrintSetup 객체 생성
                XSSFPrintSetup printSetup = sheet.getPrintSetup();
				// 인쇄 크기 설정 (letter 크기)
                printSetup.setPaperSize(XSSFPrintSetup.LETTER_PAPERSIZE);
                // 머리글 바닥글 마진 설정
                printSetup.setHeaderMargin(0);
                printSetup.setFooterMargin(0);

                printSetup.setScale((short) 100);

                Row row1 = sheet.createRow(0);
                row1.setHeightInPoints((short) 40);
                Cell row1Cell1 = row1.createCell(0);
                Cell row1Cell7 = row1.createCell(7);
                row1Cell1.setCellValue("(" + shotTag + ") " + resultDTO.getMeta().getSymbol() + " (" + koreaName + ")");
                row1Cell1.setCellStyle(row1Cell1Style);
                row1Cell7.setCellValue(String.valueOf(tradeDate.getMonth()).substring(0,3) + "/" + String.valueOf(tradeDate.getYear()).substring(2,4));
                row1Cell7.setCellStyle(row1Cell8Style);
                Row row2 = sheet.createRow(1);
                row2.setHeightInPoints((short) 21);

                Cell row2Cell1 = row2.createCell(0);
                row2Cell1.setCellValue("NBI");
                row2Cell1.setCellStyle(row2Cell1Style);

                Cell row2Cell2 = row2.createCell(1);
                row2Cell2.setCellValue(tradeDate.getMonth().toString().substring(0,3));
                row2Cell2.setCellStyle(row2Cell2to8Style);

                Cell row2Cell3 = row2.createCell(2);
                row2Cell3.setCellValue("OPEN");
                row2Cell3.setCellStyle(row2Cell2to8Style);

                Cell row2Cell4 = row2.createCell(3);
                row2Cell4.setCellValue("HIGH");
                row2Cell4.setCellStyle(row2Cell2to8Style);

                Cell row2Cell5 = row2.createCell(4);
                row2Cell5.setCellValue("LOW");
                row2Cell5.setCellStyle(row2Cell2to8Style);

                Cell row2Cell6 = row2.createCell(5);
                row2Cell6.setCellValue("CLOSE");
                row2Cell6.setCellStyle(row2Cell2to8Style);

                Cell row2Cell7 = row2.createCell(6);
                row2Cell7.setCellValue("VOLUME");
                row2Cell7.setCellStyle(row2Cell2to8Style);

                Cell row2Cell8 = row2.createCell(7);
                row2Cell8.setCellValue("CHANGE");
                row2Cell8.setCellStyle(row2Cell2to8Style);

                for (int day = 2; day < 33; day++) {
                    Row nowRow = sheet.createRow(day);
                    // 행 높이 설정
                    nowRow.setHeightInPoints((short) 21);
                    Cell cell2 = nowRow.createCell(1);
                    cell2.setCellValue(33 - day);
                    cell2.setCellStyle(dayStyle);
                    for (int colum = 0; colum < 8; colum++) {
                        if (colum != 1) {
                            Cell nowCell = nowRow.createCell(colum);
                            nowCell.setCellStyle(defaultStyle);
                        }
                    }
                }

                /*
                시트 설정값
                1500 -> 5.14
				1550 -> 5.29
				1600 -> 5.57
				1650 -> 5.71
				2560 -> 9.29
				3328 -> 12.29
				3340 -> 12.29
				2755 -> 10
				2750 -> 10
				3575 -> 13.29
				3525 -> 13
				3500 -> 13
				1675 -> 5.86
				4590 -> 17.29
				4540 -> 17
                */

                sheet.setColumnWidth(0, 3525); // 13
                sheet.setColumnWidth(1, 1675); // 5.86
                sheet.setColumnWidth(2, 2755); // 10
                sheet.setColumnWidth(3, 2755); // 10
                sheet.setColumnWidth(4, 2755); // 10
                sheet.setColumnWidth(5, 2755); // 10
                sheet.setColumnWidth(6, 4540); // 17
                sheet.setColumnWidth(7, 3500); // 13

                sheetMap.put(sheetName, sheet);
            }
        }

        for (Integer i = totalCount - 1; i >= 0; i--) {
            LocalDate tradeDate = Instant.ofEpochSecond(timestampList.get(i)).atZone(ZoneId.systemDefault()).toLocalDate();
            String sheetName = String.valueOf(tradeDate.getMonth()).substring(0, 3) + " " + String.valueOf(tradeDate.getYear()).substring(2, 4);

            // 시트내용 채우기
            XSSFSheet selectedSheet = sheetMap.get(sheetName);
            Row nowRow = null;
            for (int j = 1; j < 32; j++) {
                if (tradeDate.getDayOfMonth() == j) {
                    nowRow = selectedSheet.getRow(33-j);

                    // nbi
                    Cell nowRowCell1 = nowRow.getCell(0);
                    if (nbiMap.get(tradeDate) != null) {
                        nowRowCell1.setCellValue(nbiMap.get(tradeDate));
                        nowRowCell1.setCellStyle(cashDotTwoStyle);
                    }
                    //Cell nowRowCell2 = nowRow.getCell(1);
                    //nowRowCell2.setCellValue(String.valueOf(tradeDate.getDayOfMonth()));

                    Cell nowRowCell3 = nowRow.getCell(2);
                    nowRowCell3.setCellValue(Double.valueOf(String.format("%.2f", openList.get(i))));
                    nowRowCell3.setCellStyle(dotTwoStyle);

                    Cell nowRowCell4 = nowRow.getCell(3);
                    nowRowCell4.setCellValue(Double.valueOf(String.format("%.2f", highList.get(i))));
                    nowRowCell4.setCellStyle(dotTwoStyle);

                    Cell nowRowCell5 = nowRow.getCell(4);
                    nowRowCell5.setCellValue(Double.valueOf(String.format("%.2f", lowList.get(i))));
                    nowRowCell5.setCellStyle(dotTwoStyle);

                    Cell nowRowCell6 = nowRow.getCell(5);
                    nowRowCell6.setCellValue(Double.valueOf(String.format("%.2f", closeList.get(i))));
                    nowRowCell6.setCellStyle(dotTwoStyle);

                    Cell nowRowCell7 = nowRow.getCell(6);
                    nowRowCell7.setCellValue(volumeList.get(i));
                    nowRowCell7.setCellStyle(cashStyle);

                    Cell nowRowCell8 = nowRow.getCell(7);

                    if (i != 0) {
                        Double yesterDayClose = closeList.get(i - 1);
                        Double todayClose = closeList.get(i);
                        Double change = (todayClose - yesterDayClose) / yesterDayClose;
                        nowRowCell8.setCellValue(change);
                        nowRowCell8.setCellStyle(percentStyle);
                    }
                }
            }

        }

        log.info("excelFileCreate End");

        return xls;
    }

    private PeriodDTO getPeriods(FinanceInfo financeInfo) {
        PeriodDTO periodDTO = new PeriodDTO();
        if (financeInfo.getStartDate() != null) {
            Timestamp startPeriod = Timestamp.valueOf(financeInfo.getStartDate().atTime(0, 0, 0));
            periodDTO.setStartPeriodString(String.valueOf(startPeriod.getTime()).substring(0,10));
        } else {
            // financeApi사용해서 firstTradeDate 가져오기
            PeriodDTO tempPeriod = new PeriodDTO();
            Timestamp tempStartPeriod = Timestamp.valueOf(LocalDate.now().atTime(0, 0, 0).minusDays(7));
            Timestamp tempEndPeriod = Timestamp.valueOf(LocalDate.now().atTime(0, 0, 0));
            tempPeriod.setStartPeriodString(String.valueOf(tempStartPeriod.getTime()).substring(0,10));
            tempPeriod.setEndPeriodString(String.valueOf(tempEndPeriod.getTime()).substring(0,10));
            String url = getRequestedUrl(financeInfo.getTicker(), tempPeriod, "1d");
            String response = webClient.get()
                    .uri(url)
                    .retrieve()
                    .bodyToMono(String.class)
                    .block();

            ApiResultDTO resultDTO = null;
            try {
                JSONObject jsonObject = objectMapper.readValue(response, JSONObject.class);
                LinkedHashMap chartMap = (LinkedHashMap) jsonObject.get("chart");
                List resultList = (List) chartMap.get("result");
                LinkedHashMap resultMap = (LinkedHashMap) resultList.get(0);

                resultDTO = objectMapper.convertValue(resultMap, ApiResultDTO.class);
                log.info("*** getPeriods - resultDTO : {}", resultDTO);
            } catch (Exception e) {
                log.error("*** getPeriods - DTO변환 도중 에러가 발생했습니다.");
                e.printStackTrace();
            }

            LocalDate firstTradeDate = Instant.ofEpochSecond(resultDTO.getMeta().getFirstTradeDate()).atZone(ZoneId.systemDefault()).toLocalDate();

            // 처음 트레이드 시작일 (상장일) 로 startDate를 설정한다.
            Timestamp startPeriod = Timestamp.valueOf(firstTradeDate.atTime(0, 0, 0));
            periodDTO.setStartPeriodString(String.valueOf(startPeriod.getTime()).substring(0,10));
        }

        if (financeInfo.getEndDate() != null) {
            Timestamp endPeriod = Timestamp.valueOf(financeInfo.getEndDate().atTime(0, 0, 0));
            periodDTO.setEndPeriodString(String.valueOf(endPeriod.getTime()).substring(0,10));
        } else {
            Timestamp endPeriod = Timestamp.valueOf(LocalDate.now().atTime(0, 0, 0));
            periodDTO.setEndPeriodString(String.valueOf(endPeriod.getTime()).substring(0,10));
        }
        return periodDTO;
    }

    private String getRequestedUrl(String ticker, PeriodDTO periodSetting, String interval) {
        String fullUrl = yahooFinanceUrl + "/v8/finance/chart/" + ticker;

        LinkedMultiValueMap<String, String> map = new LinkedMultiValueMap<>();
        map.add("formatted", "true");
//        map.add("crumb", "GhYi70f3k9R");
        map.add("lang", "en-US");
        map.add("region", "US");
        map.add("includeAdjustedClose", "true");
        map.add("interval", interval);

        map.add("period1", periodSetting.getStartPeriodString());
        map.add("period2", periodSetting.getEndPeriodString());
//        map.add("events", "capitalGain%7Cdiv%7Csplit");
        map.add("useYfid", "true");
//        map.add("corsDomain", "finance.yahoo.com");

        fullUrl += "?";
        for (String key : map.keySet()) {
            fullUrl += key + "=" + map.getFirst(key) + "&";
        }

        String finalUrl = fullUrl.substring(0, fullUrl.length() - 1);
        return finalUrl;
    }

    private XSSFWorkbook setMonthData(XSSFWorkbook makedExcel, String monthDataUrl, String companyFullName, String ticker) {
        XSSFWorkbook finalExcel = makedExcel;

        String monthDataResponse = "";
        try {
            monthDataResponse = webClient.get()
                    .uri(monthDataUrl)
                    //.header("Authorization", "Bearer " + yahooFinanceApiKey)
                    .retrieve()
                    .bodyToMono(String.class)
                    .block();
        } catch (RuntimeException e) {
            // 예외 처리
            log.error("파이낸셜 api로 월별 정보를 가져오던 중 에러가 발생했습니다.");
            e.printStackTrace(); // 예외 정보 출력
            return null;
        }

        ApiResultDTO monthDataResultDTO = null;
        try {
            JSONObject jsonObject = objectMapper.readValue(monthDataResponse, JSONObject.class);
            LinkedHashMap chartMap = (LinkedHashMap) jsonObject.get("chart");
            List resultList = (List) chartMap.get("result");
            LinkedHashMap resultMap = (LinkedHashMap) resultList.get(0);

            monthDataResultDTO = objectMapper.convertValue(resultMap, ApiResultDTO.class);
            log.info("monthDataResultDTO : {}", monthDataResultDTO);
        } catch (Exception e) {
            log.error("월별 DTO변환 도중 에러가 발생했습니다.");
            e.printStackTrace();
        }

        List<Double> closeList = monthDataResultDTO.getIndicators().getQuote().get(0).getClose();
        List<Double> openList = monthDataResultDTO.getIndicators().getQuote().get(0).getOpen();
        List<Double> lowList = monthDataResultDTO.getIndicators().getQuote().get(0).getLow();
        List<Double> highList = monthDataResultDTO.getIndicators().getQuote().get(0).getHigh();

        LinkedHashMap<String, XSSFSheet> sheetMap = new LinkedHashMap<>();

        List<Integer> timestampList = monthDataResultDTO.getTimestamp();
        Integer totalCount = monthDataResultDTO.getTimestamp().size();

        LocalDate lastTradeDate = Instant.ofEpochSecond(timestampList.get(totalCount - 1)).atZone(ZoneId.systemDefault()).toLocalDate();
        Boolean even = lastTradeDate.getYear() % 2 == 0 ? true : false;
        Boolean odd = lastTradeDate.getYear() % 2 == 1 ? true : false;

        for (Integer i = totalCount - 1; i >= 0; i--) {
            LocalDate tradeDate = Instant.ofEpochSecond(timestampList.get(i)).atZone(ZoneId.systemDefault()).toLocalDate();
            if (tradeDate.getMonth().equals(LocalDate.now().getMonth()) && tradeDate.getYear() == LocalDate.now().getYear()) {
                continue;
            }
            Integer tradeYear = tradeDate.getYear();
            Boolean trageYearEven = tradeYear % 2 == 0 ? true : false;
            if (!(even && trageYearEven)) {
                continue;
            }
            String sheetName = String.valueOf(tradeDate.getYear()).substring(2,4) + "-" + String.valueOf(tradeDate.getYear()-1).substring(2,4);

            //  시트생성
            if (sheetMap.get(sheetName) == null) {
                XSSFSheet sheet = finalExcel.createSheet(sheetName);
                sheet.setMargin(PageMargin.TOP, Double.valueOf("0.6")); // 여백 위쪽1.5
                sheet.setMargin(PageMargin.LEFT, Double.valueOf("0.5")); // 여백 왼쪽1.3
                sheet.setMargin(PageMargin.RIGHT, Double.valueOf("0.5")); // 여백 오른쪽1.3
                sheet.setMargin(PageMargin.BOTTOM, Double.valueOf("0.3")); // 여백 아래쪽0.8
                sheet.setHorizontallyCenter(true); // 페이지 가운데 맞춤

                // HSSFPrintSetup 또는 XSSFPrintSetup 객체 생성
                XSSFPrintSetup printSetup = sheet.getPrintSetup();
                // 인쇄 크기 설정 (letter 크기)
                printSetup.setPaperSize(XSSFPrintSetup.LETTER_PAPERSIZE);
                // 머리글 바닥글 마진 설정
                printSetup.setHeaderMargin(0);
                printSetup.setFooterMargin(0);
                printSetup.setScale((short) 100);

                Row row1 = sheet.createRow(0);
                Cell row1Cell1 = row1.createCell(0);
                row1Cell1.setCellValue(ticker + "/ " + companyFullName);

                Row row2 = sheet.createRow(1);

                Row row3 = sheet.createRow(2);
                Cell row3Cell1 = row3.createCell(0);
                row3Cell1.setCellValue("Date");

                Cell row3Cell2 = row3.createCell(1);
                row3Cell2.setCellValue("Open");

                Cell row3Cell3 = row3.createCell(2);
                row3Cell3.setCellValue("High");

                Cell row3Cell4 = row3.createCell(3);
                row3Cell4.setCellValue("Low");

                Cell row3Cell5 = row3.createCell(4);
                row3Cell5.setCellValue("Close");

                Cell row3Cell6 = row3.createCell(5);
                row3Cell6.setCellValue("Avg Vol");

                sheetMap.put(sheetName , sheet);
            }
        }

        for (Integer i = totalCount - 1; i >= 0; i--) {
            LocalDate tradeDate = Instant.ofEpochSecond(timestampList.get(i)).atZone(ZoneId.systemDefault()).toLocalDate();
            if (tradeDate.getMonth().equals(LocalDate.now().getMonth()) && tradeDate.getYear() == LocalDate.now().getYear()) {
                continue;
            }
            String tradeYearString = String.valueOf(tradeDate.getYear()).substring(2,4);
            for (String sheetName : sheetMap.keySet()) {
                if (sheetName.contains(tradeYearString)) {
                    XSSFSheet loadSheet = sheetMap.get(sheetName);

                    Row nowRow = loadSheet.createRow(loadSheet.getLastRowNum() + 1);
                    Cell cell1 = nowRow.createCell(0);
                    cell1.setCellValue(tradeDate.format(DateTimeFormatter.ISO_DATE));

                    Cell cell2 = nowRow.createCell(1);
                    cell2.setCellValue(openList.get(i));

                    Cell cell3 = nowRow.createCell(2);
                    cell3.setCellValue(highList.get(i));

                    Cell cell4 = nowRow.createCell(3);
                    cell4.setCellValue(lowList.get(i));

                    Cell cell5 = nowRow.createCell(4);
                    cell5.setCellValue(closeList.get(i));

                    Cell cell6 = nowRow.createCell(5);
                }
            }
        }

        //monthDataResultDTO
        log.info("test");

        return finalExcel;
    }

    private String getCompanyFullName(FinanceInfo financeInfo) {
        try {
            String response = webClient.get()
                    .uri("https://finance.yahoo.com/quote/" + financeInfo.getTicker())
                    //.header("Authorization", "Bearer " + yahooFinanceApiKey)
                    .retrieve()
                    .bodyToMono(String.class)
                    .block();
            Integer start = response.indexOf("<h1 class=\"D(ib) Fz(18px)\"");
            Integer end = response.indexOf("</h1>");
            String temp = response.substring(start, end);
            String fullName = temp.substring(temp.indexOf(">")+1, temp.length()-7);
            return fullName;
        } catch (Exception e) {
            log.error("티커의 이름을 가져오는데 오류가 발생했습니다.");
            throw new RuntimeException("티커의 이름을 가져오는데 오류가 발생했습니다.");
        }
    }
}
