package com.example.yahoo_finance.entity;

import lombok.Getter;
import lombok.Setter;

import java.time.LocalDate;

@Getter
@Setter
public class RowDataDTO {
    LocalDate tradeDate;

    Double nbi;
    Double close;
    Integer volume;
    Double open;
    Double low;
    Double high;
    Double change;
}
