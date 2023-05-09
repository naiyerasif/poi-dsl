package dev.mflash.poi.dsl;

import dev.mflash.poi.dsl.utils.PoiUtils;
import dev.mflash.poi.dsl.value.PoiCellAttributes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface Poi {

	static PoiSheetQuery workbook(Workbook workbook) {
		return new PoiSheetQuery(workbook);
	}

	record PoiSheetQuery(Workbook workbook) {

		public PoiOperations sheetName(String sheetName) {
			return new PoiOperations(workbook.getSheet(sheetName));
		}
	}

	record PoiOperations(Sheet sheet) {

		public PoiCellReader read(PoiCellAttributes cellAttributes) {
			return PoiCellReader.create(() -> {
				final var row = sheet.getRow(cellAttributes.rowIndex());
				return cellAttributes.missingCellPolicy() != null ?
						row.getCell(cellAttributes.columnIndex(), cellAttributes.missingCellPolicy()) :
						row.getCell(cellAttributes.columnIndex());
			});
		}

		public PoiCellWriter write(PoiCellAttributes cellAttributes) {
			return PoiCellWriter.create(() -> {
				final var row = PoiUtils.getRow(cellAttributes.rowIndex(), sheet);
				return cellAttributes.cellType() != null ?
						PoiUtils.getCell(cellAttributes.columnIndex(), row, cellAttributes.cellType()) :
						PoiUtils.getCell(cellAttributes.columnIndex(), row);
			});
		}
	}
}
