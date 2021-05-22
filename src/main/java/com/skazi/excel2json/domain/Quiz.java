package com.skazi.excel2json.domain;

import lombok.Data;

@Data
public class Quiz {
	private String que;
	private String code;
	private String[] option;
	private int crt;
}
