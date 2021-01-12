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
        JFileChooser fileChooser = new JFileChooser("F://Schedules/Реестры/");
        fileChooser.setMultiSelectionEnabled(true);
        fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
        fileChooser.setFileHidingEnabled(false);
        int ret = fileChooser.showDialog(null, "Выбрать реестр(ы)");
        if (ret != JFileChooser.APPROVE_OPTION) {
            System.out.println("Ошибка с файлом реестра!");
        } else {

            File[] files = fileChooser.getSelectedFiles();

            for(File file: files){

                String registryName=file.getName();

                Registry registry;
                registry = new Registry();

                // Создаем книгу
                Workbook mainWB=new XSSFWorkbook();
                Sheet sheet=mainWB.createSheet("Календарный план");

                System.out.println("Выполняю " + registryName);

                try {
                    try {
                        registry = RegistryBuilder.buildFromRegistryFile("F://Schedules/Реестры/" + registryName);
                    } catch (Exception e) {
                        System.out.println("Ошибка в считывании реестра");
                    }


                    int isLift=0;
                    if (registry.getWorkNames().get(0).equals("Лифт") || registry.getWorkNames().get(0).equals("ТО")) {
                    isLift=1;
                    }

                    // Создаем шапку на основе данных реестра
                    try {
                        Head.addHead(sheet, registry.getWorkNames().get(0), registry.getMaxTerm(), isLift);
                    } catch (Exception e) {
                        System.out.println("Ошибка в формировании шапки графика");
                    }

                    // Заполняем график
                    try {
                        Addresses.addAddresses(sheet, registry, isLift);
                    } catch (Exception e) {
                        System.out.println("Ошибка в заполнении тела графика");
                    }

                    // Итоги
                    try {
                        Total.addTotal(sheet, registry, isLift);
                    } catch (Exception e) {
                        System.out.println("Ошибка в заполнении итогов");
                    }

                } catch (Exception e) {
                    System.out.println("Ошибка в " + registryName);
                    continue;
                }

                // Сохраняем файл
                try (OutputStream fileOut=new FileOutputStream("F://Schedules/Графики/" + registryName.substring(7))) {
                    mainWB.write(fileOut);
                    System.out.println("ОК!");
                } catch (IOException e) {
                    e.printStackTrace();
                }

                mainWB.close();

            }
        }
    }
}
