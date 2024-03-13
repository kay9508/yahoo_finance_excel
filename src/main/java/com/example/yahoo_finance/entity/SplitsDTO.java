package com.example.yahoo_finance.entity;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class SplitsDTO {
    private Integer date;
    private Double numerator;
    private Double denominator;
    private String splitRatio;
}
