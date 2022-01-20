package com.example.demo.util;

import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import com.example.demo.model.CategoriaProdotto;
import com.example.demo.model.Excel;
import com.example.demo.repository.ExcelRepository;


@Component
public class MatrixExcel {

	@Autowired
	ExcelRepository excelRepository;

	public List<String> matrix2Data(InputStream iStream) throws EncryptedDocumentException, IOException {

		boolean check1 = false;
		boolean check2 = false;
		boolean check3 = false;

		boolean check4 = true;
		boolean check5 = true;
		boolean check6 = true;

		List<String> error = new ArrayList<>();
		List<Excel> listExcel = new ArrayList<>();

		Workbook workbook = WorkbookFactory.create(iStream);
		Sheet firstSheet = workbook.getSheetAt(0);
		Integer righe = firstSheet.getLastRowNum();
		Integer colonne = 3;
		Iterator<Row> rowIterator = firstSheet.iterator();
		Row row;
		Cell c;

		row = rowIterator.next();
		for (int i = 0; i < colonne; i++) {
			c = row.getCell(i);
			switch (i) {
			case 0:
				try {
					String nomeProdotto = "Nome";
					String cNomeProdotto = c.getStringCellValue();
					String lowerNomeProdotto = nomeProdotto.toLowerCase();
					String lowerCNomeProdotto = cNomeProdotto.toLowerCase();
					if (lowerNomeProdotto.equals(lowerCNomeProdotto)) {
						check1 = true;
					} else {
						String errProdotto = " COLONNA ' Nome ' ASSENTE ";
						error.add(errProdotto);
					}
					break;
				} catch (Exception e) {
					String errProdotto = " COLONNA ' Nome Prodotto ' ASSENTE ";
					error.add(errProdotto);
					break;
				}

			case 1:
				try {
					String categoria = "Categoria";
					String cCategoria = c.getStringCellValue();
					String lowerCategoria = categoria.toLowerCase();
					String lowerCCategoria = cCategoria.toLowerCase();
					if (lowerCategoria.equals(lowerCCategoria)) {
						check2 = true;
					} else {
						String errCategoria = " COLONNA ' Categoria ' ASSENTE ( Consultare la lista delle categorie ) ";
						error.add(errCategoria);
					}
					break;

				} catch (Exception e) {
					String errCategoria = " COLONNA ' Categoria ' ASSENTE ( Consultare la lista delle categorie ) ";
					error.add(errCategoria);
					break;
				}

			case 2:
				try {
					String prezzo = "Prezzo";
					String cPrezzo = c.getStringCellValue();
					String lowerPrezzo = prezzo.toLowerCase();
					String lowerCPrezzo = cPrezzo.toLowerCase();
					if (lowerPrezzo.equals(lowerCPrezzo)) {
						check3 = true;
					} else {
						String errPrezzo = " COLONNA ' Prezzo ' ASSENTE ";
						error.add(errPrezzo);
					}
					break;

				} catch (Exception e) {
					String errPrezzo = " COLONNA ' Prezzo ' ASSENTE ";
					error.add(errPrezzo);
					break;

				}
			}

		}
		try {
			if (check1 && check2 && check3) {
				for (int i = 0; i < righe; i++) {
					row = rowIterator.next();
					Excel excel = new Excel();
					for (int j = 0; j < colonne; j++) {
						c = row.getCell(j);
						switch (j) {
						case 0:
							try {
								String nomeProdotto = c.getStringCellValue();
								excel.setNomeProdotto(nomeProdotto);
								break;
							} catch (Exception e) {
								error.add(" Errore Nome-Prodotto a riga " + (i + 1) );
								check4 = false;
								break;
							}

						case 1:
							try {
								String categoria = c.getStringCellValue();
								String catogoriaUP = categoria.toUpperCase();
								excel.setCategoriaProdotto(CategoriaProdotto.valueOf(catogoriaUP));
								break;
							} catch (Exception e) {
								error.add(" Errore Categoria a riga " + (i + 1) );
								check5 = false;
								break;
							}

						case 2:
							try {
								excel.setPrezzo(c.getNumericCellValue());
								break;
							} catch (Exception e) {
								error.add(" Errore Prezzo  a riga " + (i + 1) );
								check6 = false;
								break;
							}

						}

					}

					if (check4 && check5 && check6) {
						LocalDate ld = LocalDate.now();
						excel.setUploadDate(ld);
						listExcel.add(excel);
					}
				}
			}
		} catch (Exception e) {
			error.clear();
			error.add("Attenzione record assente a riga " + row.getRowNum());
		}
		if (righe < 1) {
			error.add("EXCEL VUOTO - COMPILARE I CAMPI");
			return error;
		}
		if (righe == listExcel.size()) {
			for (int i = 0; i < listExcel.size(); i++) {
				excelRepository.save(listExcel.get(i));
			}
			String x = " salvato correttamente";
			error.add(x);
			return error;

		} else {
			error.add("ATTENZIONE COMPILAZIONE EXCEL NON CORRETTA - CARICAMENTO FALLITO");
			return error;

		}

	}
}
