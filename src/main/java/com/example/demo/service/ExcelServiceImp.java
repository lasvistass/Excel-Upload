package com.example.demo.service;

import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.example.demo.model.Excel;
import com.example.demo.repository.ExcelRepository;



@Service
public class ExcelServiceImp implements ExcelService {

	@Autowired
	ExcelRepository excelRepo;

	@Override
	public Excel salva(Excel excel) {
		return excelRepo.save(excel);
	}

	@Override
	public List<Excel> listFile() {
		return excelRepo.findAll();
	}

	@Override
	public List<Excel> arrayToList(Excel[] lista) {
		List<Excel> excelList = Arrays.asList(lista);
		for (int i = 0; i < excelList.size(); i++) {
			excelRepo.save(excelList.get(i));
		}
		return excelList;
	}

	@Override
	public List<Excel> betweenDates(LocalDate start, LocalDate end) {
		List<Excel> lista=excelRepo.findBetweenDates(start,end);
		return lista;
	}

}
