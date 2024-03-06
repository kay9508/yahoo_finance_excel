package com.example.yahoo_finance.service;

import com.example.yahoo_finance.entity.ApiResultDTO;
import com.example.yahoo_finance.entity.FinanceInfo;
import com.example.yahoo_finance.entity.IndicatorsDTO;
import com.example.yahoo_finance.entity.MetaDTO;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Service;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.web.reactive.function.client.ExchangeStrategies;
import org.springframework.web.reactive.function.client.WebClient;

import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
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
        String url = yahooFinanceUrl + "/v8/finance/chart/" + financeInfo.getTicker();
        Timestamp period1 = Timestamp.valueOf(financeInfo.getStartDate().atTime(0, 0, 0));
        Timestamp period2 = Timestamp.valueOf(financeInfo.getEndDate().atTime(0, 0, 0));

        LinkedMultiValueMap<String, String> map = new LinkedMultiValueMap<>();
        map.add("formatted", "true");
        map.add("crumb", "GhYi70f3k9R");
        map.add("lang", "en-US");
        map.add("region", "US");
        map.add("includeAdjustedClose", "true");
        map.add("interval", "1d");

        map.add("period1", String.valueOf(period1.getTime()).substring(0,10));
        map.add("period2", String.valueOf(period2.getTime()).substring(0,10));
        map.add("events", "capitalGain%7Cdiv%7Csplit");
        map.add("useYfid", "true");
        map.add("corsDomain", "finance.yahoo.com");

        url += "?";
        for (String key : map.keySet()) {
            url += key + "=" + map.getFirst(key) + "&";
        }

        String finalUrl = url.substring(0, url.length() - 1);
        String response = "";
        try {

            response = webClient.get()
                    .uri(finalUrl)
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

        ApiResultDTO resultDTO = null;
        try {
            JSONObject jsonObject = objectMapper.readValue(response, JSONObject.class);
            LinkedHashMap chartMap = (LinkedHashMap) jsonObject.get("chart");
            List resultList = (List) chartMap.get("result");
            LinkedHashMap resultMap = (LinkedHashMap) resultList.get(0);

            resultDTO = objectMapper.convertValue(resultMap, ApiResultDTO.class);
            log.info("resultDTO : {}", resultDTO);
        } catch (Exception e) {
            log.error("DTO변환 도중 에러가 발생했습니다.");
            e.printStackTrace();
        }

        LocalDate firstTradeDate = Instant.ofEpochSecond(resultDTO.getMeta().getFirstTradeDate()).atZone(ZoneId.systemDefault()).toLocalDate();
        if (!firstTradeDate.equals(financeInfo.getStartDate())) {
            log.warn("(!주의!)첫 거래일이 입력된 일자와 다릅니다.");
        }

        XSSFWorkbook makedExcel = excelFileCreate(resultDTO, financeInfo.getShotTag(), financeInfo.getKoreaName());
        log.info("끝");
        return makedExcel;
    }

    private XSSFWorkbook excelFileCreate(ApiResultDTO resultDTO, String shotTag, String koreaName) {
        shotTag = shotTag == null ? "" : shotTag;
        koreaName = koreaName == null ? "" : koreaName;

        List<Double> closeList = resultDTO.getIndicators().getQuote().get(0).getClose();
        List<Integer> volumeList = resultDTO.getIndicators().getQuote().get(0).getVolume();
        List<Double> openList = resultDTO.getIndicators().getQuote().get(0).getOpen();
        List<Double> lowList = resultDTO.getIndicators().getQuote().get(0).getLow();
        List<Double> highList = resultDTO.getIndicators().getQuote().get(0).getHigh();

        // DecimalFormat decFormat = new DecimalFormat("###,###");
        // 엑셀 파일 생성
        XSSFWorkbook xls = new XSSFWorkbook(); // .xlsx

        // 스타일
        CellStyle cashStyle = xls.createCellStyle();
        cashStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0"));

        CellStyle percentStyle = xls.createCellStyle();
        percentStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00%"));

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
                Row row1 = sheet.createRow(0);
                Cell row1Cell1 = row1.createCell(0);
                Cell row1Cell7 = row1.createCell(7);
                row1Cell1.setCellValue("(" + shotTag + ") " + resultDTO.getMeta().getSymbol() + " (" + koreaName + ")");
                row1Cell7.setCellValue(String.valueOf(tradeDate.getMonth()).substring(0,3) + "/" + String.valueOf(tradeDate.getYear()).substring(2,4));
                Row row2 = sheet.createRow(1);
                Cell row2Cell1 = row2.createCell(0);
                row2Cell1.setCellValue("NBI");
                Cell row2Cell2 = row2.createCell(1);
                row2Cell2.setCellValue("DEC");
                Cell row2Cell3 = row2.createCell(2);
                row2Cell3.setCellValue("OPEN");
                Cell row2Cell4 = row2.createCell(3);
                row2Cell4.setCellValue("HIGH");
                Cell row2Cell5 = row2.createCell(4);
                row2Cell5.setCellValue("LOW");
                Cell row2Cell6 = row2.createCell(5);
                row2Cell6.setCellValue("CLOSE");
                Cell row2Cell7 = row2.createCell(6);
                row2Cell7.setCellValue("VOLUME");
                Cell row2Cell8 = row2.createCell(7);
                row2Cell8.setCellValue("CHANGE");

                for (int day = 2; day < 33; day++) {
                    Row nowRow = sheet.createRow(day);
                    Cell cell2 = nowRow.createCell(1);
                    cell2.setCellValue(33 - day);
                }

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

                    // 1번 항목은 다른 api에서 가져와야하는 내용
                    Cell nowRowCell1 = nowRow.createCell(0);
                    //Cell nowRowCell2 = nowRow.createCell(1);
                    //nowRowCell2.setCellValue(String.valueOf(tradeDate.getDayOfMonth()));

                    Cell nowRowCell3 = nowRow.createCell(2);
                    nowRowCell3.setCellValue(Double.valueOf(String.format("%.2f", openList.get(i))));

                    Cell nowRowCell4 = nowRow.createCell(3);
                    nowRowCell4.setCellValue(Double.valueOf(String.format("%.2f", highList.get(i))));

                    Cell nowRowCell5 = nowRow.createCell(4);
                    nowRowCell5.setCellValue(Double.valueOf(String.format("%.2f", lowList.get(i))));

                    Cell nowRowCell6 = nowRow.createCell(5);
                    nowRowCell6.setCellValue(Double.valueOf(String.format("%.2f", closeList.get(i))));

                    Cell nowRowCell7 = nowRow.createCell(6);
                    nowRowCell7.setCellValue(volumeList.get(i));
                    nowRowCell7.setCellStyle(cashStyle);

                    Cell nowRowCell8 = nowRow.createCell(7);

                    if (i != 0) {
                        Double yesterDayClose = closeList.get(i - 1);
                        Double todayClose = closeList.get(i);
                         Double change = (todayClose - yesterDayClose) / yesterDayClose * 100;
                        String chageValue = String.format("%.2f",change) + "%";
                        nowRowCell8.setCellValue(chageValue);
                    }
                }
            }

        }

        log.info("excelFileCreate End");

        return xls;
    }
}
