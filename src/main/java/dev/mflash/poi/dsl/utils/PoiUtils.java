package dev.mflash.poi.dsl.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellUtil;

public interface PoiUtils {

	static Row getRow(int rowIndex, Sheet sheet) {
		return CellUtil.getRow(rowIndex, sheet);
	}

	static Cell getCell(int columnIndex, Row row) {
		Cell cell = row.getCell(columnIndex);
		if (cell == null) {
			cell = row.createCell(columnIndex);
		}

		return cell;
	}

	static Cell getCell(int columnIndex, Row row, CellType cellType) {
		Cell cell = row.getCell(columnIndex);
		if (cell == null) {
			cell = row.createCell(columnIndex, cellType);
		}

		return cell;
	}
}
