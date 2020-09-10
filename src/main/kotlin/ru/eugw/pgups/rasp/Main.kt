package ru.eugw.pgups.rasp

import com.google.gson.GsonBuilder
import com.google.gson.JsonArray
import com.google.gson.JsonObject
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.net.URL

fun main() {
    println("Enter course")
    val course = readLine()
    println("Enter group")
    val group = readLine()
    URL("https://rasp.pgups.ru/files/xls_files/$course/$group.xlsx").openStream().copyTo(File("schedules", "$group.xlsx").outputStream())
    val groupParser = GroupParser(group!!)
    groupParser.parseWeek()
}

class GroupParser(private val group: String) {

    private val wb = XSSFWorkbook(File("schedules", "$group.xlsx"))

    fun parseWeek() {
        val weekObject = JsonObject()
        val map = mapOf(
            "2" to Pair(6, 15),
            "3" to Pair(16, 25),
            "4" to Pair(26, 35),
            "5" to Pair(36, 45),
            "6" to Pair(46, 55),
            "7" to Pair(56, 57),
            "2e" to Pair(6, 15),
            "3e" to Pair(16, 25),
            "4e" to Pair(26, 35),
            "5e" to Pair(36, 45),
            "6e" to Pair(46, 55),
            "7e" to Pair(56, 57)
        )
        map.forEach {
            weekObject.add(it.key, parseDOW(it.key.length > 1, it.value.first, it.value.second))
        }
        println("Finished")
        File("parsed_schedules", "$group.json").writeText(GsonBuilder().setPrettyPrinting()
            .create().toJson(weekObject))
    }

    private fun getMergedRegion(c: Cell): CellRangeAddress? {
        val s = c.row.sheet
        s.mergedRegions.forEach {
            if (it.isInRange(c.rowIndex, c.columnIndex))
                return it
        }
        return null
    }

    private fun parseDOW(even: Boolean, start: Int, end: Int): JsonArray {
        val jsonArray = JsonArray()
        val sheet = wb.getSheetAt(0)
        sheet.forEachIndexed { iRow, row ->
            row.forEachIndexed { iCell, cell ->
                if (iRow in start..end && iCell == 2) {
                    val jsonLesson = JsonObject()
                    val reg = getMergedRegion(cell)
                    val cellString = if (reg == null) cell.stringCellValue else sheet.getRow(reg.firstRow)
                        .getCell(reg.firstColumn).stringCellValue
                    val arrayList = ArrayList<String>()
                    cellString.replace("  ", "\n").split("\n").forEach {
                        if (it.isNotBlank())
                            arrayList.add(it)
                    }
                    jsonLesson.addProperty("lesson", with(arrayList) {
                        if (this.size > 1)
                            this.subList(0, 2).toString()
                        else
                            this.toString()
                    }.replace("[", "").replace("]", ""))
                    jsonLesson.addProperty("time", sheet.getRow(iRow - 1 * iRow % 2)
                        .getCell(iCell - 1).stringCellValue)
                    jsonLesson.addProperty("cabinet", cellString.substringAfter("ауд. "))
                    if (jsonLesson["lesson"].asString.isNotBlank())
                        if (even && (iRow + 1) % 2 == 0 || !even && (iRow + 1) % 2 != 0) {
                            jsonArray.add(jsonLesson)
                        }
                }
            }
        }
        return jsonArray
    }

}