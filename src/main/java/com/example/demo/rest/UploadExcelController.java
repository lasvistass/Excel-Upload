package com.example.demo.rest;


import java.util.ArrayList;
import java.util.List;


import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import com.example.demo.repository.ExcelRepository;
import com.example.demo.util.ConverterExcel;
import com.example.demo.util.MatrixExcel;



@RestController
@RequestMapping("/api")
public class UploadExcelController {

	@Autowired
	ConverterExcel converterExcel;

	@Autowired
	ExcelRepository excelRepository;

	@Autowired
	MatrixExcel matrix;


	@PostMapping("/upload/excel")
	public List<String> handleFileUploadExcel(@RequestParam("file") MultipartFile mFile) {
		try {
			String filename = mFile.getOriginalFilename();
			if(filename.equals("")){
				List<String> list0 = new ArrayList<>();
				list0.add(" NESSUN FILE SELEZIONATO - SELEZIONARE FORMATO EXCEL ( .xls )");
				return list0;
			}
			String x = " salvato correttamente";
			List<String> list = matrix.matrix2Data(mFile.getInputStream());
			String y = list.get(0);
			if(x.equals(y)) {
				list.add(0, filename);
				return list;
			}
			return list;
		}catch(Exception e) {
			List<String> list2 = new ArrayList<>();
			String filename = mFile.getOriginalFilename();
			list2.add("FILE ' " + filename + " ' NON SUPPORTATO - SELEZIONARE FORMATO EXCEL ( .xls )");
			return list2;
		}
	}
}
