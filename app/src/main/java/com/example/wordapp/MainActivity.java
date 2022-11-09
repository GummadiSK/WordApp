package com.example.wordapp;

import androidx.appcompat.app.AppCompatActivity;
import org.apache.poi.util.IOUtils;
import android.annotation.SuppressLint;
import android.os.Bundle;

import android.Manifest;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.view.View;
import android.widget.EditText;
import android.widget.Toast;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import java.io.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class MainActivity extends AppCompatActivity {

    private EditText editTextInput;
    private File filePath = null;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        ActivityCompat.requestPermissions(this, new String[]{Manifest.permission.READ_EXTERNAL_STORAGE,
                        Manifest.permission.WRITE_EXTERNAL_STORAGE},
                PackageManager.PERMISSION_GRANTED);

        editTextInput = findViewById(R.id.editTextTextPersonName);
        filePath = new File(getExternalFilesDir(null), "Test.docx");

        try {
            if (!filePath.exists()){
                filePath.createNewFile();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void buttonCreate(View view){
        try {

            XWPFDocument xwpfDocument = new XWPFDocument();
            XWPFParagraph xwpfParagraph = xwpfDocument.createParagraph();
            XWPFRun xwpfRun = xwpfParagraph.createRun();

        //    xwpfRun.setText(editTextInput.getText().toString());

            XWPFRun line1 = xwpfParagraph.createRun();
            line1.setBold(true);
            line1.setText("JNTUH UNIVERSITY COLLEGE OF ENGINEERING SULTANPUR");
            line1.setFontSize(18);
            line1.addBreak();


            XWPFRun line2 = xwpfParagraph.createRun();
            line2.setBold(true);
            line2.setText("\n\t SULTANPUR(V),PULKAL(M),SANGAREDDY DIST");
            line2.setFontSize(18);
            line2.addBreak();

            XWPFRun line3 = xwpfParagraph.createRun();
            line3.setBold(true);
            line3.setText("\n\tII B.TECH- I SEM-(R18) II MID EXAMINATIONS,FEB-2022");
            line3.setFontSize(18);
            line3.addBreak();

            XWPFRun line4 = xwpfParagraph.createRun();
            line4.setText("\n Subject: OPERATING SYSTEMS \t \t \t\t\tTIME: 60 Minutes");
            line4.setFontSize(12);
            line4.setBold(true);
            line4.addBreak();

            XWPFRun line5 = xwpfParagraph.createRun();
            line5.setText("\n Branch : CSE \t \t \t\t\t\t\t\t Marks : 10");
            line5.setFontSize(12);
            line5.setBold(true);
            line5.addBreak();

            XWPFRun line6 = xwpfParagraph.createRun();
            line6.setText("\n Date of Exam : 15-02-2022 (AN)");
            line6.setFontSize(12);
            line6.setBold(true);
            line6.addBreak();

            XWPFRun line7 = xwpfParagraph.createRun();
            line7.setText("\n__________________________________________________________________________________");
            line7.setBold(true);
            line7.addBreak();

            XWPFRun line8 = xwpfParagraph.createRun();
            line8.setText("\n \t\t\t  Note: Answer any TWO of the following questions     ");
            line8.setFontSize(12);
            line8.setBold(true);
            line8.setItalic(true);
            line8.addBreak();

            XWPFRun line9 = xwpfParagraph.createRun();
            String que1 = "1. Explain the following operational functions of access channels \n   a) FOCC \n   b) RECC\n";
            line9.setText(que1);
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            xwpfDocument.write(fileOutputStream);

           /* File image = new File("C:\\Users\\sruja\\Downloads\\jntu.jpg");
            FileInputStream imageData = new FileInputStream(image);
            int imageType = XWPFDocument.PICTURE_TYPE_JPEG;
            String imageFileName = image.getName();
            int width=300;
            int height=350;
            xwpfRun.addPicture(imageData,imageType,imageFileName, Units.toEMU(width),Units.toEMU(height));
            xwpfDocument.write(fileOutputStream); */

            if(fileOutputStream!=null){
                fileOutputStream.flush();
                fileOutputStream.close();
            }

            xwpfDocument.close();
            Toast.makeText(MainActivity.this,"Successful",Toast.LENGTH_SHORT).show();
        }
        catch (Exception e){
            e.printStackTrace();
            editTextInput.setText("fail");
        }
    }
}