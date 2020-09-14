/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package optiim;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

/**
 *
 * @author Mert
 */
public class OPTiiM {

    public static Person[] read(String input) throws IOException {

        Person P[] = new Person[100];

        String inputFile = input;
        File inputWorkbook = new File(inputFile);
        Workbook w;
        
        try {
            w = Workbook.getWorkbook(inputWorkbook);

            Sheet sheet = w.getSheet(0);
           

            for (int k = 0; k < sheet.getRows(); k++) {
                P[k] = new Person();
            }
            
            
            for (int j = 0; j < sheet.getRows(); j++) {
                for (int i = 0; i < sheet.getColumns(); i++) {
                    Cell cell = sheet.getCell(i, j);
                    CellType type = cell.getType();

                    if (type == CellType.LABEL) {

                        System.out.print("  " + cell.getContents());
                    } else if (type == CellType.NUMBER) {

                        System.out.print("  " + cell.getContents());
                    } else {

                        System.out.print("  null");
                    }

                    if (sheet.getColumns() == 8) {

                        switch (i) {
                            case 0:
                                P[j].Ad = cell.getContents();
                                break;
                            case 1:
                                P[j].Soyad = cell.getContents();
                                break;
                            case 2:
                                P[j].Dogum_tarihi = cell.getContents();
                                break;
                            case 3:
                                P[j].Dogum_yeri = cell.getContents();
                                break;
                            case 4:
                                P[j].Mail = cell.getContents();
                                break;
                            case 5:
                                P[j].Telefon = cell.getContents();
                                break;
                            case 6:
                                P[j].Calisma_durumu = cell.getContents();
                                break;
                            case 7:
                                P[j].Universite = cell.getContents();
                                break;

                        }
                    } else {
                        switch (i) {
                            case 0:
                                 P[j].Ad = cell.getContents();
                                break;
                            case 1:
                                P[j].Soyad = cell.getContents();
                                break;
                            case 2:
                               P[j].Dogum_tarihi = cell.getContents();
                                break;
                            case 3:
                                 P[j].Mail = cell.getContents();
                                break;
                            case 4:
                                P[j].Telefon = cell.getContents();
                                break;
                            case 5:
                                P[j].Durum = cell.getContents();
                                break;

                        }
                    }

                }
                System.out.println(" ");
            }

        } catch (BiffException e) {

            e.printStackTrace();
        }


        return P;
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, WriteException {
        //ReadExcel read = new ReadExcel();
        
        String str=null;

        File file = new File("log.txt");
        if (!file.exists()) {
            file.createNewFile();
        }

        FileWriter fileWriter = new FileWriter(file, false);
        BufferedWriter bWriter = new BufferedWriter(fileWriter);
      
     
        
        
        Person Rapor[] = new Person[20];
        for (int k = 0; k < 20; k++) {
                Rapor[k] = new Person();
        }
        int tampon=0;
        
        String inputFile = "C:\\Users\\Mert\\Documents\\NetBeansProjects\\OPTiiM\\InputData\\Excel2.xls";
        Person P[] = new Person[20];
        P = read(inputFile);
        
        System.out.println("-----------------------------------");
        
        String inputFile2 = "C:\\Users\\Mert\\Documents\\NetBeansProjects\\OPTiiM\\InputData\\Excel1.xls";
        Person P2[] = new Person[20];
        P2 = read(inputFile2);
        
        
        System.out.println("-----------------------------------");
        for(int i=1;i<12;i++){
            for(int j=0;j<12;j++){
                
                
                
                if(P[i].Ad.equals(P2[j].Ad) && P[i].Soyad.equals(P2[j].Soyad)){
                    //System.out.println(P[i].Ad+P[i].Soyad);
                    
                   
                    if((P[i].Telefon.equals(P2[j].Telefon) && !P[i].Telefon.isEmpty() )||(P[i].Mail.equals(P2[j].Mail)&& !P[i].Mail.isEmpty()))
                    {
                   
                        
                        Rapor[tampon].Ad=P[i].Ad;
                        Rapor[tampon].Soyad=P[i].Soyad;
                        Rapor[tampon].Dogum_tarihi=P[i].Dogum_tarihi;
                        Rapor[tampon].Dogum_yeri=P[i].Dogum_yeri;
                        Rapor[tampon].Mail=P[i].Mail;
                        Rapor[tampon].Telefon=P[i].Telefon;
                        Rapor[tampon].Durum=P2[j].Durum;
                        Rapor[tampon].Calisma_durumu=P[i].Calisma_durumu;
                        Rapor[tampon].Universite=P[i].Universite;

                        tampon++;
                    }else {
                       // System.out.println(P[i].Ad+" "+P[i].Soyad+" ın maili ve ya telefonu yok");
                        System.out.println(P[i].Ad+" "+P[i].Soyad +" Maili ya da telefonu uyuşmamaktadır veya ikiside yoktur");
                         
                        str=P[i].Ad+" "+P[i].Soyad +" Maili ya da telefonu uyuşmamaktadır veya ikiside yoktur";
                        bWriter.write(str);
                        bWriter.newLine();
                    }
                }
                
            }
        }
        bWriter.close();
        System.out.println(P[6].Telefon+" 5 "+P2[6].Telefon);
        WriteExcel write = new WriteExcel();
        write.setOutputFile("C:\\Users\\Mert\\Documents\\NetBeansProjects\\OPTiiM\\Report\\Report.xls");
        write.write(Rapor);
        

    }

}
