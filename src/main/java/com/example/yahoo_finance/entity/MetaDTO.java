package com.example.yahoo_finance.entity;

import lombok.Getter;
import lombok.Setter;

import java.util.LinkedHashMap;
import java.util.List;

@Getter
@Setter
public class MetaDTO {

private String currency;
private String symbol;
private String exchangeName;
private String instrumentType;
private Integer firstTradeDate;
private Integer regularMarketTime;
private Boolean hasPrePostMarketData;
private Integer gmtoffset;
private String timezone;
private String exchangeTimezoneName;
private Double regularMarketPrice;
private Double chartPreviousClose;
private Integer priceHint;
private LinkedHashMap currentTradingPeriod;
private String dataGranularity;
private String range;
private List validRanges;
}
