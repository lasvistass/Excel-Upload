package com.example.demo.rest;

import java.io.IOException;
import java.sql.Date;
import java.util.ArrayList;
import java.util.List;
import javax.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.example.demo.model.Excel;
import com.example.demo.service.ExcelService;
import com.example.demo.util.ExcelExport;

@RestController
@RequestMapping(value = "/api/download")
public class DownloadExcelController {

	@Autowired
	ExcelService excelService;

	@GetMapping(value = "/dates")
	public void exportBetweenDates(@RequestParam("start") String start, @RequestParam("end") String end,
			HttpServletResponse resp) throws IOException {

		resp.setContentType("application/octet-stream");
		String headerKey = "Content-Disposition";
		String headervalue = "attachment; filename=prodotti.xlsx";
		resp.setHeader(headerKey, headervalue);

		List<Excel> listExcel = excelService.betweenDates(Date.valueOf(start).toLocalDate(),
				Date.valueOf(end).toLocalDate());
		ExcelExport exp = new ExcelExport(listExcel);
		exp.export(resp);

	}

	@GetMapping(value = "/template")
	public void template(HttpServletResponse resp) throws IOException {
		List<Excel> listExcel = new ArrayList<Excel>();
		resp.setContentType("application/octet-stream");
		String headerKey = "Content-Disposition";
		String headervalue = "attachment; filename=prodotti.xlsx";
		resp.setHeader(headerKey, headervalue);
		ExcelExport exp = new ExcelExport(listExcel);
		exp.export(resp);

	}
}
