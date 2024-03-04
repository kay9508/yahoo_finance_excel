package com.example.yahoo_finance.entity;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class IndicatorsDTO {
    private List<QuoteDTO> quote;
    private List<AdjcloseDTO> adjclose;
}
