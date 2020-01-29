package com.fkr.schedule;

import com.fkr.schedule.domain.Registry;
import com.fkr.schedule.domain.RegistryBuilder;
import com.fkr.schedule.scheduleparts.Addresses;
import com.fkr.schedule.scheduleparts.Head;
import com.fkr.schedule.scheduleparts.Total;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class ScheduleGenerator {
    public static void main(String[] args) throws IOException {

        // Выбор файла реестра
        JFileChooser fileChooser = new JFileChooser("E://Реестры/");
        fileChooser.setMultiSelectionEnabled(true);
        fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        fileChooser.setFileHidingEnabled(false);
        int ret = fileChooser.showDialog(null, "Выбрать реестр(ы)");
        if (ret != JFileChooser.APPROVE_OPTION) {
            System.out.println("Ошибка с файлом реестра(ов)!");
        } else {

            File[] files = fileChooser.getSelectedFiles();

            for(File file: files){

                String registryName=file.getName();
                Registry registry=RegistryBuilder.buildFromRegistryFile("E://Реестры/" + registryName);

                int isLift=0;
                if (registry.getWorkNames().get(0).equals("Лифт") || registry.getWorkNames().get(0).equals("ТО")) {
                    isLift=1;
                }

                // Создаем книгу
                Workbook mainWB=new XSSFWorkbook();
                Sheet sheet=mainWB.createSheet("Календарный план");

                // Создаем шапку на основе данных реестра
                Head.addHead(sheet, registry.getWorkNames().get(0), registry.getMaxTerm(), isLift);

                // Заполняем график
                Addresses.addAddresses(sheet, registry, isLift);

                // Итоги
                Total.addTotal(sheet, registry, isLift);

                // Сохраняем файл
                try (OutputStream fileOut=new FileOutputStream("E://Графики/" + registryName)) {
                    mainWB.write(fileOut);
                } catch (IOException e) {
                    e.printStackTrace();
                }

                mainWB.close();

            }
        }
    }
}
