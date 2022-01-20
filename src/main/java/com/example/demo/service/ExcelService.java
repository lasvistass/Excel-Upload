package com.example.demo.service;

import java.time.LocalDate;
import java.util.List;

import com.example.demo.model.Excel;


public interface ExcelService {

	public Excel salva(Excel excel);

	public List<Excel> listFile();

	public List<Excel> arrayToList(Excel[] lista);

	public List<Excel> betweenDates(LocalDate start, LocalDate end);

}
