package com.example.create_excel_sheet

import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.os.Environment
import android.widget.Button
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellUtil.createCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileNotFoundException
import java.io.FileOutputStream
import java.io.IOException

class MainActivity : AppCompatActivity() {
    lateinit var button:Button
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        button = findViewById(R.id.button)
        button.setOnClickListener {
            val ourWB = createWorkbook()
            createExcelFile(ourWB)
        }
    }
}
private fun createWorkbook(): Workbook {
    val ourWorkbook = XSSFWorkbook()
    val sheet: Sheet = ourWorkbook.createSheet("statSheet")
    ourWorkbook.createSheet("testSheet")
    addData(sheet)

    return ourWorkbook
}
private fun addData(sheet: Sheet) {

    val row1 = sheet.createRow(0)
    val row2 = sheet.createRow(1)
    val row3 = sheet.createRow(2)
    val row4 = sheet.createRow(3)

    createCell(row1, 0, "Name")
    createCell(row1, 1, "Score")

    createCell(row2, 0, "Sam")
    createCell(row2, 1, "45")

    createCell(row3, 0, "Dam")
    createCell(row3, 1, "60")

    createCell(row4, 0, "Kohli")
    createCell(row4, 1, "0")

}
private fun createExcelFile(ourWorkbook: Workbook) {

    val ourAppFileDirectory = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
    if (ourAppFileDirectory != null && !ourAppFileDirectory.exists()) {
        ourAppFileDirectory.mkdirs()
    }

    val excelFile = File(ourAppFileDirectory, "sam.xlsx")
    try {
        val fileOut = FileOutputStream(excelFile)
        ourWorkbook.write(fileOut)
        fileOut.close()
    } catch (e: FileNotFoundException) {
        e.printStackTrace()
    } catch (e: IOException) {
        e.printStackTrace()
    }
}