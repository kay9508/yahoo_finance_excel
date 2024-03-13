package com.example.yahoo_finance.entity;

import lombok.Getter;
import lombok.Setter;

import java.util.LinkedHashMap;

@Getter
@Setter
public class EventDTO {
    private LinkedHashMap<String, DividendsDTO> dividends;
    private LinkedHashMap<String, SplitsDTO> splits;
}
