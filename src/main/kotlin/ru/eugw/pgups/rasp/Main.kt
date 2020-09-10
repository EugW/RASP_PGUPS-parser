package ru.eugw.pgups.rasp

import com.google.gson.GsonBuilder
import com.google.gson.JsonArray
import com.google.gson.JsonObject
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.net.URL
import kotlin.math.ceil

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
    private val sheet = wb.getSheetAt(0)
    private var cellOffset: Int? = null

    fun parseWeek() {
        val weekObject = JsonObject()
        val map = mapOf(
            "2" to Pair(cellOffset() + 1, 15),
            "3" to Pair(cellOffset() + 11, 25),
            "4" to Pair(cellOffset() + 21, 35),
            "5" to Pair(cellOffset() + 31, 45),
            "6" to Pair(cellOffset() + 41, 55),
            "7" to Pair(cellOffset() + 51, 57),
            "2e" to Pair(cellOffset() + 1, 15),
            "3e" to Pair(cellOffset() + 11, 25),
            "4e" to Pair(cellOffset() + 21, 35),
            "5e" to Pair(cellOffset() + 31, 45),
            "6e" to Pair(cellOffset() + 41, 55),
            "7e" to Pair(cellOffset() + 51, 57),
        )
        map.forEach {
            weekObject.add(it.key, parseDOW(it.key.length > 1, it.value.first, it.value.second))
        }
        println("Finished")
        File("parsed_schedules", "$group.json").writeText(GsonBuilder().setPrettyPrinting()
            .create().toJson(weekObject))
    }

    private fun cellOffset(): Int {
        if (cellOffset != null)
            return cellOffset!!
        sheet.forEachIndexed { indexR, row ->
            row.forEach { cell ->
                if (cell.stringCellValue == group)
                    return indexR
            }
        }
        return 0
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
                    val timeString = (if (cellOffset() % 2 != 0)
                            sheet.getRow(iRow - if (iRow % 2 != 0) 1 else 0)
                                    else
                            sheet.getRow(iRow - if (iRow % 2 == 0) 1 else 0))
                            .getCell(iCell - 1).stringCellValue
                    val str = StringBuilder(timeString.split("-")[0])
                    val fnh = StringBuilder(timeString.split("-")[1])
                    str.insert(str.length - ceil(str.length / 2.0).toInt(), ":")
                    str.split(":").forEachIndexed { index, s ->
                        if (s.length < 2)
                            str.insert(index * 2 + index, "0")
                    }
                    fnh.insert(fnh.length - ceil(fnh.length / 2.0).toInt(), ":")
                    fnh.split(":").forEachIndexed { index, s ->
                        if (s.length < 2)
                            fnh.insert(index * 2 + index, "0")
                    }
                    jsonLesson.addProperty("time", "$str-$fnh")
                    jsonLesson.addProperty("cabinet", cellString.substringAfter("ауд. "))
                    if (jsonLesson["lesson"].asString.isNotBlank())
                        if (cellOffset() % 2 != 0) {
                            if (even && (iRow + 1) % 2 == 0 || !even && (iRow + 1) % 2 != 0)
                                jsonArray.add(jsonLesson)
                        } else {
                            if (even && (iRow + 1) % 2 != 0 || !even && (iRow + 1) % 2 == 0)
                                jsonArray.add(jsonLesson)
                        }
                }
            }
        }
        return jsonArray
    }

}