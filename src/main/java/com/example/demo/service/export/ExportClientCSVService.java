package com.example.demo.service.export;

import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.Period;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.example.demo.entity.Client;
import com.example.demo.repository.ClientRepository;

@Service
public class ExportClientCSVService {

	
	@Autowired
    private ClientRepository clientRepository;

    public void exportAll(PrintWriter writer) {
        
    	//génération d'un fichier CSV exemples avec 3 colonnes
        writer.println("Nom;Prenom;Age");

        LocalDate today = LocalDate.now();
        
        List<Client> allClients = clientRepository.findAll();
        for (Client client : allClients) {
			writer.println(client.getNom() + ";" + client.getPrenom() + ";" + calculerAge(client.getDateNaissance(), today));
        }
    }
    
    // Calcul de l'age du client
    public int calculerAge(LocalDate DateNaissance,LocalDate today) {
    		    return Period.between(DateNaissance, today).getYears();

    }
}
