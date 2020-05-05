package com.example.demo.service.export;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.example.demo.entity.Facture;
import com.example.demo.repository.FactureRepository;

@Service
public class ExportFactureExcelService {
	
	@Autowired
	private FactureRepository factureRepository;
	
	
	public void exportAll(OutputStream outputStream) throws IOException {
		
		Workbook workbook = new XSSFWorkbook();
		
		List<Facture> allFacture = factureRepository.findAll();
		
		Sheet sheet = workbook.createSheet("Factures");
		
		Row row0 = sheet.createRow(0);
        Cell cell0 = row0.createCell(0);
        cell0.setCellValue("Nom");
        int cellNum = 1;
        int i = 1;
        for(Facture facture: allFacture) {
            Cell cell = row0.createCell(cellNum++);
            cell.setCellValue(facture.getClient().getNom());
            sheet.autoSizeColumn(i);
            i++;
        }
        
        Row row1 = sheet.createRow(1);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("Prénom");
        cellNum = 1;
        i = 1;
        for(Facture facture: allFacture) {
            Cell cell = row1.createCell(cellNum++);
            cell.setCellValue(facture.getClient().getPrenom());
            sheet.autoSizeColumn(i);
            i++;
        }
        
        Row row2 = sheet.createRow(2);
        Cell cell2 = row2.createCell(0);
        cell2.setCellValue("Date de Naissance");
        cellNum = 1;
        i = 1;
        for(Facture facture: allFacture) {
            Cell cell = row2.createCell(cellNum++);
            cell.setCellValue(facture.getClient().getDateNaissance().toString());
            sheet.autoSizeColumn(i);
            i++;
        }
       
        sheet.autoSizeColumn(0);
        
        workbook.write(outputStream);
        workbook.close();
        
	}

}
