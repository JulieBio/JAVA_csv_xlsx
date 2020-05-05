package com.example.demo.service.export;

import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.time.Period;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.example.demo.entity.Client;
import com.example.demo.repository.ClientRepository;

@Service
public class ExportClientExcelService {

	@Autowired
    private ClientRepository clientRepository;

	private static String[] colonnes = {"Nom", "prenom", "Age"};
	
	//private static List<Client> clients =  new ArrayList<>();
	
	
    public void exportAll(OutputStream outputStream) throws IOException {
        
    	Workbook workbook = new XSSFWorkbook();
    	
    	LocalDate today = LocalDate.now();
    	
    	List<Client> allClients = clientRepository.findAll();
    	
    	// Création de sheet
    	Sheet sheet = workbook.createSheet("Clients");
        
        
        // Création d'une Row
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        
		// Style pour les headers du tableau -> rouge
		CellStyle styleH = workbook.createCellStyle();
        Font pinkFont = workbook.createFont();
        pinkFont.setColor(IndexedColors.PINK.getIndex());
        pinkFont.setBold(true);
        pinkFont.setFontHeightInPoints((short) 14);
        styleH.setFont(pinkFont);
        styleH.setBorderBottom(BorderStyle.MEDIUM);
        styleH.setBorderLeft(BorderStyle.MEDIUM);
        styleH.setBorderRight(BorderStyle.MEDIUM);
        styleH.setBorderTop(BorderStyle.MEDIUM);
        styleH.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        styleH.setLeftBorderColor(IndexedColors.BLUE.getIndex());
        styleH.setRightBorderColor(IndexedColors.BLUE.getIndex());
        styleH.setTopBorderColor(IndexedColors.BLUE.getIndex());

        
		// Création des cellules header
        for(int i = 0; i < colonnes.length; i++) {
            Cell cellH = row.createCell(i);
            cellH.setCellValue(colonnes[i]);
            cellH.setCellStyle(styleH);       
        }
        
        // Création Cell CellStyle
        CellStyle style = workbook.createCellStyle();
        Font blackFont = workbook.createFont();
        blackFont.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(blackFont);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        style.setLeftBorderColor(IndexedColors.BLUE.getIndex());
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());
        style.setTopBorderColor(IndexedColors.BLUE.getIndex());
		
        // Création des lignes et cellules avec les données les clients
        int rowNum = 1;
        for(Client client: allClients) {
            Row rowC = sheet.createRow(rowNum++);

            Cell cell0 = rowC.createCell(0);
            cell0.setCellValue(client.getNom());
            cell0.setCellStyle(style);
            
            Cell cell1 = rowC.createCell(1);
            cell1.setCellValue(client.getPrenom());
            cell1.setCellStyle(style);
            
            Cell cell2 = rowC.createCell(2);
            cell2.setCellValue(calculerAge(client.getDateNaissance(), today) + " ans");
            cell2.setCellStyle(style);
        }
        
        // Redimensionner les colonnes en fonction du contenu
        for(int i = 0; i < colonnes.length; i++) {
            sheet.autoSizeColumn(i);
        }
        
        workbook.write(outputStream);
        workbook.close();
        
    }
    
    // Calcul de l'age du client
    public int calculerAge(LocalDate DateNaissance,LocalDate today) {
    		    return Period.between(DateNaissance, today).getYears();
    }

    
}
