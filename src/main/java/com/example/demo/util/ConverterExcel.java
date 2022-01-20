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
import org.apache.xmlbeans.impl.jam.mutable.MField;
import org.springframework.beans.factory.annotation.Autowired;

import org.springframework.stereotype.Component;

import com.example.demo.model.CategoriaProdotto;
import com.example.demo.model.Excel;
import com.example.demo.repository.ExcelRepository;
import com.fasterxml.jackson.databind.exc.InvalidFormatException;



@Component
public class ConverterExcel {

	@Autowired
	ExcelRepository excelRepository;

	public List<String> Excel2Data(InputStream iStream)
			throws EncryptedDocumentException, InvalidFormatException, IOException {


		boolean check1 = false;
		boolean check2 = false;
		boolean check3 = false;


		boolean check4 = true;
		boolean check5 = true;
		boolean check6 = true;

		List<Excel> listaExcel = new ArrayList<>();
		List<String> error = new ArrayList<>();
		Integer righe = null;

		try {
			Workbook workbook = WorkbookFactory.create(iStream);
			Sheet firstSheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = firstSheet.iterator();
			righe = firstSheet.getLastRowNum();

			Row row = rowIterator.next();
			Iterator<Cell> cell = row.cellIterator();
			while (cell.hasNext()) {

				Cell c = cell.next();
				int columnIndex = c.getColumnIndex();
				switch (columnIndex) {
				case 0:
					try {
						String nomeProdotto = "Nome Prodotto";
						String cNomeProdotto = c.getStringCellValue();
						String lowerNomeProdotto = nomeProdotto.toLowerCase();
						String lowerCNomeProdotto = cNomeProdotto.toLowerCase();
						if (lowerNomeProdotto.equals(lowerCNomeProdotto)) {
							check1 = true;
						}else {
							String errProdotto = "*** COLONNA ' Nome Prodotto ' ASSENTE ***";
							error.add(errProdotto);
						}
						break;
					}catch(Exception e) {
						String errProdotto = "*** COLONNA ' Nome Prodotto ' ASSENTE ***";
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
							String errCategoria = "*** COLONNA ' Categoria ' ASSENTE ( Consultare la lista delle categorie ) ***";
							error.add(errCategoria);
						}
						break;

					}catch(Exception e) {
						String errCategoria = "*** COLONNA ' Categoria ' ASSENTE ( Consultare la lista delle categorie ) ***";
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
							String errPrezzo = "*** COLONNA ' Prezzo ' ASSENTE ***";
							error.add(errPrezzo);
						}
						break;

					}catch(Exception e) {
						String errPrezzo = "*** COLONNA ' Prezzo ' ASSENTE ***";
						error.add(errPrezzo);
						break;

					}

				}
			}


			try {
				if (check1 && check2 && check3) {
					while (rowIterator.hasNext()) {
						Excel excel = new Excel();
						Row nextRow = rowIterator.next();
						Iterator<Cell> cellIterator = nextRow.cellIterator();
						while (cellIterator.hasNext()) {

							Cell nextCell = cellIterator.next();
							int columnIndex = nextCell.getColumnIndex();
							switch (columnIndex) {
							case 0:
								try {
									String nomeProdotto = nextCell.getStringCellValue();
									excel.setNomeProdotto(nomeProdotto);
									break;
								} catch (Exception e) {
									String x = String.valueOf(nextRow.getRowNum());
									error.add("*** Errore Nome-Prodotto a riga " + x + " ***  ");
									check4 = false;
									break;
								}

							case 1:
								try {
									String categoria = nextCell.getStringCellValue();
									String catogoriaUP = categoria.toUpperCase();
									excel.setCategoriaProdotto(CategoriaProdotto.valueOf(catogoriaUP));
									break;
								} catch (Exception e) {
									String x = String.valueOf(nextRow.getRowNum());
									error.add("*** Errore Categoria a riga " + x + " ***  ");
									check5 = false;
									break;
								}

							case 2:
								try {
									excel.setPrezzo(nextCell.getNumericCellValue());
									break;
								} catch (Exception e) {
									String x = String.valueOf(nextRow.getRowNum());
									error.add("*** Errore Prezzo  a riga " + x + " ***  ");
									check5 = false;
									break;
								}

							}


						}

						try {
							if (check4 && check5 && check6) {
								LocalDate ld = LocalDate.now();
								excel.setUploadDate(ld);
								listaExcel.add(excel);
							}

						} catch (Exception e) {
							error.add("record non salvato");
						}

					}

				}
			} catch (Exception e) {
				error.add("compilare bene");

			}

			workbook.close();

		} catch (Exception e) {

			e.printStackTrace();
		}


		if(righe < 1) {
			error.add("EXCEL VUOTO - COMPILARE I CAMPI");
			return error;
		}

		if (righe == listaExcel.size()) {
			for (int i = 0; i < listaExcel.size(); i++) {
				excelRepository.save(listaExcel.get(i));
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
