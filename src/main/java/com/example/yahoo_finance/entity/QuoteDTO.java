package com.example.yahoo_finance.entity;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class QuoteDTO {
    private List<Double> close;
    private List<Integer> volume;
    private List<Double> open;
    private List<Double> low;
    private List<Double> high;
}
