package com.example.yahoo_finance.entity;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class ApiResultDTO {
    private MetaDTO meta;
    private List<Integer> timestamp;
    private EventDTO events;
    private IndicatorsDTO indicators;
}
