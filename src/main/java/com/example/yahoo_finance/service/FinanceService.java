package com.example.yahoo_finance.service;

import com.example.yahoo_finance.entity.*;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.json.simple.JSONObject;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPrintOptions;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.web.reactive.function.client.WebClient;

import java.sql.Date;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

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
        } catch (Exception e) {
            log.error("DTO변환 도중 에러가 발생했습니다.");
            e.printStackTrace();
        }

        log.info("첫 거래일 : " + Instant.ofEpochSecond(dayDataResultDTO.getMeta().getFirstTradeDate()).atZone(ZoneId.systemDefault()).toLocalDate());
        /*LocalDate firstTradeDate = Instant.ofEpochSecond(dayDataResultDTO.getMeta().getFirstTradeDate()).atZone(ZoneId.systemDefault()).toLocalDate();
        if (!firstTradeDate.equals(financeInfo.getStartDate())) {
            log.warn("(!주의!)첫 거래일이 입력된 일자와 다릅니다.");
        }*/

        //TODO 최초 거래일이 이상하게 나오는 종목이 뭐 하나 있었는데 확인해야함
        periodSetting.setStartPeriodString(dayDataResultDTO.getMeta().getFirstTradeDate().toString());
        String monthDataUrl = getRequestedUrl(financeInfo.getTicker(), periodSetting, "1mo");

        XSSFWorkbook makedExcel = excelCreateAndSetDayData(dayDataResultDTO, financeInfo.getShotTag(), financeInfo.getKoreaName(), nbiMap);
        XSSFWorkbook addedExcel = setMonthData(makedExcel, monthDataUrl, companyFullName, financeInfo.getTicker());
        log.info("끝");
        return addedExcel;
    }

    private XSSFWorkbook excelCreateAndSetDayData(ApiResultDTO dayDataResultDTO, String shotTag, String koreaName, LinkedHashMap<LocalDate, Double> nbiMap) {
        shotTag = shotTag == null ? "" : shotTag;
        koreaName = koreaName == null ? "" : koreaName;

        List<Double> closeList = dayDataResultDTO.getIndicators().getQuote().get(0).getClose();
        List<Integer> volumeList = dayDataResultDTO.getIndicators().getQuote().get(0).getVolume();
        List<Double> openList = dayDataResultDTO.getIndicators().getQuote().get(0).getOpen();
        List<Double> lowList = dayDataResultDTO.getIndicators().getQuote().get(0).getLow();
        List<Double> highList = dayDataResultDTO.getIndicators().getQuote().get(0).getHigh();

        LinkedHashMap<LocalDate, DividendsDTO> dividendsMap = null;
        if (dayDataResultDTO.getEvents() != null && dayDataResultDTO.getEvents().getDividends() != null) {
            dividendsMap = new LinkedHashMap<>();
            for (String s : dayDataResultDTO.getEvents().getDividends().keySet()) {
                LocalDate date = Instant.ofEpochSecond(dayDataResultDTO.getEvents().getDividends().get(s).getDate()).atZone(ZoneId.systemDefault()).toLocalDate();
                dividendsMap.put(date, dayDataResultDTO.getEvents().getDividends().get(s));
            }
        }

        LinkedHashMap<LocalDate, SplitsDTO> splitsMap = null;
        if (dayDataResultDTO.getEvents() != null && dayDataResultDTO.getEvents().getSplits() != null) {
            splitsMap = new LinkedHashMap<>();
            for (String s : dayDataResultDTO.getEvents().getSplits().keySet()) {
                LocalDate date = Instant.ofEpochSecond(dayDataResultDTO.getEvents().getSplits().get(s).getDate()).atZone(ZoneId.systemDefault()).toLocalDate();
                splitsMap.put(date, dayDataResultDTO.getEvents().getSplits().get(s));
            }
        }

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
        row1Cell1Style.setAlignment(HorizontalAlignment.LEFT);

        // H-1에 사용될 스타일
        CellStyle row1Cell8Style = xls.createCellStyle();
        Font row1Cell8Font = xls.createFont();
        row1Cell8Font.setFontName("Arial");
        row1Cell8Font.setBold(true);
        row1Cell8Font.setFontHeightInPoints((short) 14);
        row1Cell8Style.setFont(row1Cell8Font);
        row1Cell8Style.setAlignment(HorizontalAlignment.RIGHT);

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
        row2Cell1Style.setAlignment(HorizontalAlignment.CENTER);
        row2Cell1Style.setVerticalAlignment(VerticalAlignment.CENTER);

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
        row2Cell2to8Style.setAlignment(HorizontalAlignment.CENTER);
        row2Cell2to8Style.setVerticalAlignment(VerticalAlignment.CENTER);

        // VOLUME에 사용될 스타일
        CellStyle cashStyle = xls.createCellStyle();
        cashStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0"));
        cashStyle.setFont(defaultFont);
        cashStyle.setBorderRight(BorderStyle.THIN);
        cashStyle.setBorderBottom(BorderStyle.THIN);
        cashStyle.setAlignment(HorizontalAlignment.RIGHT);

        // CHANGE에 사용될 스타일
        CellStyle percentStyle = xls.createCellStyle();
        percentStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00%"));
        percentStyle.setFont(defaultFont);
        percentStyle.setBorderRight(BorderStyle.THIN);
        percentStyle.setBorderBottom(BorderStyle.THIN);
        percentStyle.setAlignment(HorizontalAlignment.RIGHT);

        // NBI에 사용될 스타일
        CellStyle cashDotTwoStyle = xls.createCellStyle();
        cashDotTwoStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
        cashDotTwoStyle.setFont(defaultFont);
        cashDotTwoStyle.setBorderLeft(BorderStyle.THIN);
        cashDotTwoStyle.setBorderBottom(BorderStyle.THIN);
        cashDotTwoStyle.setAlignment(HorizontalAlignment.RIGHT);

        // 나머지 Data에 사용될 스타일
        CellStyle dotTwoStyle = xls.createCellStyle();
        dotTwoStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        dotTwoStyle.setFont(defaultFont);
        dotTwoStyle.setBorderRight(BorderStyle.THIN);
        dotTwoStyle.setBorderBottom(BorderStyle.THIN);
        dotTwoStyle.setAlignment(HorizontalAlignment.RIGHT);

        // 기본 스타일
        CellStyle defaultStyle = xls.createCellStyle();
        defaultStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        defaultStyle.setFont(defaultFont);
        defaultStyle.setBorderTop(BorderStyle.THIN);
        defaultStyle.setBorderLeft(BorderStyle.THIN);
        defaultStyle.setBorderRight(BorderStyle.THIN);
        defaultStyle.setBorderBottom(BorderStyle.THIN);
        defaultStyle.setAlignment(HorizontalAlignment.RIGHT);

        // 분할 배당 data부분에 사용할 스타일
        CellStyle dividendAndSplitDateStyle = xls.createCellStyle();
        dividendAndSplitDateStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        dividendAndSplitDateStyle.setFont(defaultFont);
        dividendAndSplitDateStyle.setBorderTop(BorderStyle.THIN);
        dividendAndSplitDateStyle.setBorderLeft(BorderStyle.THIN);
        dividendAndSplitDateStyle.setBorderRight(BorderStyle.THIN);
        dividendAndSplitDateStyle.setBorderBottom(BorderStyle.THIN);
        dividendAndSplitDateStyle.setAlignment(HorizontalAlignment.CENTER);
        dividendAndSplitDateStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        dividendAndSplitDateStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 분할 배당 일자에 사용될 스타일
        CellStyle dividendAndSplitDayStyle = xls.createCellStyle();
        dividendAndSplitDayStyle.setFont(defaultFont);
        dividendAndSplitDayStyle.setBorderLeft(BorderStyle.DOUBLE);
        dividendAndSplitDayStyle.setBorderRight(BorderStyle.THIN);
        dividendAndSplitDayStyle.setBorderBottom(BorderStyle.THIN);
        dividendAndSplitDayStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        dividendAndSplitDayStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 일자에 사용될 스타일
        CellStyle dayStyle = xls.createCellStyle();
        dayStyle.setFont(defaultFont);
        dayStyle.setBorderLeft(BorderStyle.DOUBLE);
        dayStyle.setBorderRight(BorderStyle.THIN);
        dayStyle.setBorderBottom(BorderStyle.THIN);

        // 내림차순으로 정렬 (이거 하면 데이터 다꼬임)
        // Collections.sort(resultDTO.getTimestamp(), Collections.reverseOrder());

        LinkedHashMap<String, List<RowDataDTO>> sheetData = new LinkedHashMap<>();
        LinkedHashMap<String, XSSFSheet> sheetMap = new LinkedHashMap<>();

        List<Integer> timestampList = dayDataResultDTO.getTimestamp();
        Integer totalCount = dayDataResultDTO.getTimestamp().size();
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

                sheet.setDefaultRowHeight((short) 420);
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
                row1Cell1.setCellValue("(" + shotTag + ") " + dayDataResultDTO.getMeta().getSymbol() + " (" + koreaName + ")");
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

                /*
                시트 설정값
                1500 -> 5.14
				1550 -> 5.29
				1600 -> 5.57
				1650 -> 5.71
				1675 -> 5.86
				2560 -> 9.29
				2750 -> 10
				2755 -> 10
				3328 -> 12.29
				3340 -> 12.29
				3500 -> 13
				3525 -> 13
				3575 -> 13.29
				3725 -> 13.86
				3750 -> 14
				3775 -> 14
				3800 -> 14.14
				4540 -> 17
				4620 -> 17.29
				4720 -> 17.71
				4770 -> 17.86
				4830 -> 18.14
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

            List<RowDataDTO> rowDataList = new ArrayList<>();
            if (sheetData.get(sheetName) != null) {
                rowDataList = sheetData.get(sheetName);
            }
            RowDataDTO rowData = new RowDataDTO();
            rowData.setTradeDate(tradeDate);

            rowData.setNbi(nbiMap.get(tradeDate));
            rowData.setClose(closeList.get(i));
            rowData.setVolume(volumeList.get(i));
            rowData.setOpen(openList.get(i));
            rowData.setLow(lowList.get(i));
            rowData.setHigh(highList.get(i));
            if (i != 0) {
                Double yesterDayClose = closeList.get(i - 1);
                Double todayClose = closeList.get(i);
                Double change = (todayClose - yesterDayClose) / yesterDayClose;
                rowData.setChange(change);
            }

            rowDataList.add(rowData);

            sheetData.put(sheetName, rowDataList);
        }

        for (String sheetName : sheetMap.keySet()) {
            XSSFSheet selectedSheet = sheetMap.get(sheetName);
            Integer month = getMonthOfAbbreviation(selectedSheet.getSheetName().substring(0,3));
            Integer year = 2000 + Integer.valueOf(selectedSheet.getSheetName().substring(4,6));
            int additionalRowStart = 2;
            List<RowDataDTO> rowDataList = sheetData.get(sheetName);
            for(int day = 31; day > 0; day--) {
                LocalDate targetDate = null;
                try {
                    targetDate = LocalDate.of(year, month, day);
                } catch (Exception e) {
                    Row nowRow = selectedSheet.createRow(additionalRowStart);
                    nowRow.setHeightInPoints((short) 21);
                    additionalRowStart = additionalRowStart + 1;
                    for (int cell = 0; cell < 8; cell++) {
                        Cell nowCell = nowRow.createCell(cell);
                        if (cell == 1) {
                            nowCell.setCellValue(day);
                            nowCell.setCellStyle(dayStyle);
                        } else {
                            nowCell.setCellStyle(defaultStyle);
                        }
                    }
                    continue;
                }
                Boolean exist = false;
                for (RowDataDTO rowData : rowDataList) {
                    if (rowData.getTradeDate().equals(targetDate)) {
                        Row nowRow = selectedSheet.createRow(additionalRowStart);
                        nowRow.setHeightInPoints((short) 21);
                        additionalRowStart = additionalRowStart + 1;
                        for (int cell = 0; cell < 8; cell++) {
                            Cell nowCell = nowRow.createCell(cell);
                            switch (cell) {
                                case 0 :
                                    nowCell.setCellValue(rowData.getNbi());
                                    nowCell.setCellStyle(cashDotTwoStyle);
                                    break;
                                case 1 :
                                    nowCell.setCellValue(day);
                                    nowCell.setCellStyle(dayStyle);
                                    break;
                                case 2 :
                                    nowCell.setCellValue(rowData.getOpen());
                                    nowCell.setCellStyle(dotTwoStyle);
                                    break;
                                case 3 :
                                    nowCell.setCellValue(rowData.getHigh());
                                    nowCell.setCellStyle(dotTwoStyle);
                                    break;
                                case 4 :
                                    nowCell.setCellValue(rowData.getLow());
                                    nowCell.setCellStyle(dotTwoStyle);
                                    break;
                                case 5 :
                                    nowCell.setCellValue(rowData.getClose());
                                    nowCell.setCellStyle(dotTwoStyle);
                                    break;
                                case 6 :
                                    nowCell.setCellValue(rowData.getVolume());
                                    nowCell.setCellStyle(cashStyle);
                                    break;
                                case 7 :
                                    if (rowData.getChange() != null) {
                                        nowCell.setCellValue(rowData.getChange());
                                        nowCell.setCellStyle(percentStyle);
                                    } else {
                                        nowCell.setCellStyle(defaultStyle);
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }

                        exist = true;

                    }
                }
                if (!exist) {
                    Row nowRow = selectedSheet.createRow(additionalRowStart);
                    nowRow.setHeightInPoints((short) 21);
                    additionalRowStart = additionalRowStart + 1;
                    for (int cell = 0; cell < 8; cell++) {
                        Cell nowCell = nowRow.createCell(cell);
                        if (cell == 1) {
                            nowCell.setCellValue(day);
                            nowCell.setCellStyle(dayStyle);
                        } else {
                            nowCell.setCellStyle(defaultStyle);
                        }
                    }
                }

                // 분할 데이터
                if (splitsMap != null && splitsMap.get(targetDate) != null) {
                    Row nowRow = selectedSheet.createRow(additionalRowStart);
                    nowRow.setHeightInPoints((short) 21);
                    additionalRowStart = additionalRowStart + 1;
                    for (int cell = 0; cell < 8; cell++) {
                        Cell nowCell = nowRow.createCell(cell);
                        if (cell == 1) {
                            nowCell.setCellValue(day);
                            nowCell.setCellStyle(dividendAndSplitDayStyle);
                        } else if (cell == 2) {
                            nowCell.setCellValue(splitsMap.get(targetDate).getSplitRatio() + " Stock Split");
                            nowCell.setCellStyle(dividendAndSplitDateStyle);
                        } else {
                            nowCell.setCellStyle(dividendAndSplitDateStyle);
                        }
                    }

                    // 병합할 셀 범위 지정
                    CellRangeAddress region = new CellRangeAddress(
                            additionalRowStart - 1, // firstRow
                            additionalRowStart - 1, // lastRow
                            2, // firstCol
                            7 // lastCol
                    );
                    selectedSheet.addMergedRegion(region);
                }

                //  배당데이터
                if (dividendsMap != null && dividendsMap.get(targetDate) != null) {
                    Row nowRow = selectedSheet.createRow(additionalRowStart);
                    nowRow.setHeightInPoints((short) 21);
                    additionalRowStart = additionalRowStart + 1;
                    for (int cell = 0; cell < 8; cell++) {
                        Cell nowCell = nowRow.createCell(cell);
                        if (cell == 1) {
                            nowCell.setCellValue(day);
                            nowCell.setCellStyle(dividendAndSplitDayStyle);
                        } else if (cell == 2) {
                            nowCell.setCellValue(String.format("%.2f", dividendsMap.get(targetDate).getAmount()) + " Dividend");
                            nowCell.setCellStyle(dividendAndSplitDateStyle);
                        } else {
                            nowCell.setCellStyle(dividendAndSplitDateStyle);
                        }
                    }

                    // 병합할 셀 범위 지정
                    CellRangeAddress region = new CellRangeAddress(
                            additionalRowStart - 1, // firstRow
                            additionalRowStart - 1, // lastRow
                            2, // firstCol
                            7 // lastCol
                    );
                    selectedSheet.addMergedRegion(region);
                }
            }
        }

        log.info("일별데이서 생성 완료");

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
            Timestamp endPeriod = Timestamp.valueOf(LocalDate.now().minusDays(1).atTime(0, 0, 0));
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
        map.add("events", "capitalGain%7Cdiv%7Csplit");
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

        Font defaultFont = finalExcel.createFont();
        defaultFont.setFontName("Arial");
        defaultFont.setFontHeightInPoints((short) 15);

        // 기본 스타일
        CellStyle defaultFontStyle = finalExcel.createCellStyle();
        defaultFontStyle.setFont(defaultFont);
        defaultFontStyle.setAlignment(HorizontalAlignment.RIGHT);

        // A-1에 사용될 스타일
        CellStyle row1Cell1Style = finalExcel.createCellStyle();
        Font row1Cell1Font = finalExcel.createFont();
        row1Cell1Font.setFontName("Arial");
        row1Cell1Font.setBold(true);
        row1Cell1Font.setFontHeightInPoints((short) 22);
        row1Cell1Style.setFont(row1Cell1Font);
        row1Cell1Style.setAlignment(HorizontalAlignment.LEFT);

        // row3에 사용될 스타일
        CellStyle subTitleStyle = finalExcel.createCellStyle();
        subTitleStyle.setFont(defaultFont);
        subTitleStyle.setAlignment(HorizontalAlignment.RIGHT);

        // 나머지 Data에 사용될 스타일
        CellStyle dotTwoStyle = finalExcel.createCellStyle();
        dotTwoStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        dotTwoStyle.setFont(defaultFont);
        dotTwoStyle.setAlignment(HorizontalAlignment.RIGHT);

        // 분할 배당 날짜 부분에 사용할 스타일
        CellStyle dividendAndSplitDateStyle = finalExcel.createCellStyle();
        dividendAndSplitDateStyle.setFont(defaultFont);
        dividendAndSplitDateStyle.setAlignment(HorizontalAlignment.RIGHT);
        dividendAndSplitDateStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        dividendAndSplitDateStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 분할 배당 데이터 부분에 사용할 스타일
        CellStyle dividendAndSplitDataStyle = finalExcel.createCellStyle();
        dividendAndSplitDataStyle.setFont(defaultFont);
        dividendAndSplitDataStyle.setAlignment(HorizontalAlignment.CENTER);
        dividendAndSplitDataStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        dividendAndSplitDataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

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
        } catch (Exception e) {
            log.error("월별 DTO변환 도중 에러가 발생했습니다.");
            e.printStackTrace();
        }

        List<Double> closeList = monthDataResultDTO.getIndicators().getQuote().get(0).getClose();
        List<Double> openList = monthDataResultDTO.getIndicators().getQuote().get(0).getOpen();
        List<Double> lowList = monthDataResultDTO.getIndicators().getQuote().get(0).getLow();
        List<Double> highList = monthDataResultDTO.getIndicators().getQuote().get(0).getHigh();

        LinkedHashMap<LocalDate, DividendsDTO> dividendsMap = null;
        if (monthDataResultDTO.getEvents() != null && monthDataResultDTO.getEvents().getDividends() != null) {
            dividendsMap = new LinkedHashMap<>();
            for (String s : monthDataResultDTO.getEvents().getDividends().keySet()) {
                LocalDate date = Instant.ofEpochSecond(monthDataResultDTO.getEvents().getDividends().get(s).getDate()).atZone(ZoneId.systemDefault()).toLocalDate();
                dividendsMap.put(date, monthDataResultDTO.getEvents().getDividends().get(s));
            }
        }

        LinkedHashMap<LocalDate, SplitsDTO> splitsMap = null;
        if (monthDataResultDTO.getEvents() != null && monthDataResultDTO.getEvents().getSplits() != null) {
            splitsMap = new LinkedHashMap<>();
            for (String s : monthDataResultDTO.getEvents().getSplits().keySet()) {
                LocalDate date = Instant.ofEpochSecond(monthDataResultDTO.getEvents().getSplits().get(s).getDate()).atZone(ZoneId.systemDefault()).toLocalDate();
                splitsMap.put(date, monthDataResultDTO.getEvents().getSplits().get(s));
            }
        }

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
                row1.setHeightInPoints((short) 37);
                Cell row1Cell1 = row1.createCell(0);
                row1Cell1.setCellValue(ticker + "/ " + companyFullName);
                row1Cell1.setCellStyle(row1Cell1Style);

                Row row2 = sheet.createRow(1);
                row2.setHeightInPoints((short) 22);

                Row row3 = sheet.createRow(2);
                Cell row3Cell1 = row3.createCell(0);
                row3Cell1.setCellValue("Date");
                row3Cell1.setCellStyle(subTitleStyle);

                Cell row3Cell2 = row3.createCell(1);
                row3Cell2.setCellValue("Open");
                row3Cell2.setCellStyle(subTitleStyle);

                Cell row3Cell3 = row3.createCell(2);
                row3Cell3.setCellValue("High");
                row3Cell3.setCellStyle(subTitleStyle);

                Cell row3Cell4 = row3.createCell(3);
                row3Cell4.setCellValue("Low");
                row3Cell4.setCellStyle(subTitleStyle);

                Cell row3Cell5 = row3.createCell(4);
                row3Cell5.setCellValue("Close");
                row3Cell5.setCellStyle(subTitleStyle);

                Cell row3Cell6 = row3.createCell(5);
                row3Cell6.setCellValue("Avg Vol");
                row3Cell6.setCellStyle(subTitleStyle);

                /*
                시트 설정값
                1500 -> 5.14
				1550 -> 5.29
				1600 -> 5.57
				1650 -> 5.71
				1675 -> 5.86
				2560 -> 9.29
				2750 -> 10
				2755 -> 10
				3328 -> 12.29
				3340 -> 12.29
				3500 -> 13
				3525 -> 13
				3575 -> 13.29
				3725 -> 13.86
				3750 -> 14
				3775 -> 14
				3800 -> 14.14
				4540 -> 17
				4620 -> 17.29
				4720 -> 17.71
				4770 -> 17.86
				4830 -> 18.14
				4880 -> 18.29
                */

                sheet.setColumnWidth(0, 4660); // 17.43
                sheet.setColumnWidth(1, 3750); // 14
                sheet.setColumnWidth(2, 3750); // 14
                sheet.setColumnWidth(3, 3750); // 14
                sheet.setColumnWidth(4, 3750); // 14
                sheet.setColumnWidth(5, 4880); // 18.29

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
                    SimpleDateFormat dateFormat = new SimpleDateFormat("MMM d, yyyy");

                    // 분할 데이터
                    if (splitsMap != null) {
                        for (LocalDate splitDate : splitsMap.keySet()) {
                            if (tradeDate.getYear() == splitDate.getYear() && tradeDate.getMonth() == splitDate.getMonth()) {
                                Row nowRow = loadSheet.createRow(loadSheet.getLastRowNum() + 1);
                                Cell cell1 = nowRow.createCell(0);
                                Date tempSimpleDateDate = Date.valueOf(splitDate);
                                cell1.setCellValue(dateFormat.format(tempSimpleDateDate));
                                cell1.setCellStyle(dividendAndSplitDateStyle);

                                Cell cell2 = nowRow.createCell(1);
                                cell2.setCellValue(splitsMap.get(splitDate).getSplitRatio() + " Stock Split");
                                cell2.setCellStyle(dividendAndSplitDataStyle);

                                Cell cell3 = nowRow.createCell(2);
                                Cell cell4 = nowRow.createCell(3);
                                Cell cell5 = nowRow.createCell(4);
                                Cell cell6 = nowRow.createCell(5);
                                // 병합할 셀 범위 지정
                                CellRangeAddress region = new CellRangeAddress(
                                        loadSheet.getLastRowNum(), // firstRow
                                        loadSheet.getLastRowNum(), // lastRow
                                        1, // firstCol
                                        5// lastCol
                                );
                                loadSheet.addMergedRegion(region);
                            }
                        }
                    }

                    // 배당 데이터
                    if (dividendsMap != null) {
                        for (LocalDate dividendsDate : dividendsMap.keySet()) {
                            if (tradeDate.getYear() == dividendsDate.getYear() && tradeDate.getMonth() == dividendsDate.getMonth()) {
                                Row nowRow = loadSheet.createRow(loadSheet.getLastRowNum() + 1);
                                Cell cell1 = nowRow.createCell(0);
                                Date tempSimpleDateDate = Date.valueOf(dividendsDate);
                                cell1.setCellValue(dateFormat.format(tempSimpleDateDate));
                                cell1.setCellStyle(dividendAndSplitDateStyle);

                                Cell cell2 = nowRow.createCell(1);
                                cell2.setCellValue(String.format("%.2f", dividendsMap.get(dividendsDate).getAmount()) + " Dividend");
                                cell2.setCellStyle(dividendAndSplitDataStyle);

                                Cell cell3 = nowRow.createCell(2);
                                Cell cell4 = nowRow.createCell(3);
                                Cell cell5 = nowRow.createCell(4);
                                Cell cell6 = nowRow.createCell(5);
                                // 병합할 셀 범위 지정
                                CellRangeAddress region = new CellRangeAddress(
                                        loadSheet.getLastRowNum(), // firstRow
                                        loadSheet.getLastRowNum(), // lastRow
                                        1, // firstCol
                                        5// lastCol
                                );
                                loadSheet.addMergedRegion(region);
                            }
                        }
                    }

                    Row nowRow = loadSheet.createRow(loadSheet.getLastRowNum() + 1);
                    nowRow.setHeightInPoints((short) 22);
                    Cell cell1 = nowRow.createCell(0);

                    Date tempSimpleDateDate = Date.valueOf(tradeDate);
                    cell1.setCellValue(dateFormat.format(tempSimpleDateDate));
                    cell1.setCellStyle(defaultFontStyle);

                    Cell cell2 = nowRow.createCell(1);
                    cell2.setCellValue(openList.get(i));
                    cell2.setCellStyle(dotTwoStyle);

                    Cell cell3 = nowRow.createCell(2);
                    cell3.setCellValue(highList.get(i));
                    cell3.setCellStyle(dotTwoStyle);

                    Cell cell4 = nowRow.createCell(3);
                    cell4.setCellValue(lowList.get(i));
                    cell4.setCellStyle(dotTwoStyle);

                    Cell cell5 = nowRow.createCell(4);
                    cell5.setCellValue(closeList.get(i));
                    cell5.setCellStyle(dotTwoStyle);

                    Cell cell6 = nowRow.createCell(5);
                }
            }
        }

        log.info("월별데이터 생성 완료");

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
            String fullName = temp.substring(temp.indexOf(">")+1, temp.length()-6);
            return fullName;
        } catch (Exception e) {
            log.error("티커의 이름을 가져오는데 오류가 발생했습니다.");
            throw new RuntimeException("티커의 이름을 가져오는데 오류가 발생했습니다.");
        }
    }

    private Integer getMonthOfAbbreviation(String abbreviation) {
        Integer month = 0;
        switch (abbreviation) {
            case "DEC" : month = 12; break;
            case "NOV" : month = 11; break;
            case "OCT" : month = 10; break;
            case "SEP" : month = 9; break;
            case "AUG" : month = 8; break;
            case "JUL" : month = 7; break;
            case "JUN" : month = 6; break;
            case "MAY" : month = 5; break;
            case "APR" : month = 4; break;
            case "MAR" : month = 3; break;
            case "FEB" : month = 2; break;
            case "JAN" : month = 1; break;
            default:month = 0; break;
        }
        if (month == 0) {
            log.error("잘못된 월을 가지고 있는 데이터가 포함되어 있습니다.");
            throw new RuntimeException("잘못된 월을 가지고 있는 데이터가 포함되어 있습니다.");
        }
        return month;
    }
}
