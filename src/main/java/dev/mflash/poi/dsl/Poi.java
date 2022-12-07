package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Poi {

	private final Workbook workbook;
	private Sheet sheet;

	private Poi(Workbook workbook) {
		this.workbook = workbook;
	}

	public static PoiSheetSetup workbook(Workbook workbook) {
		return new PoiSheetSetup(new Poi(workbook));
	}

	public static final class PoiSheetSetup {
		private final Poi poi;

		private PoiSheetSetup(Poi poi) {
			this.poi = poi;
		}

		public PoiOperationSetup sheetName(String sheetName) {
			poi.sheet = poi.workbook.getSheet(sheetName);
			return new PoiOperationSetup(poi);
		}
	}

	public static final class PoiOperationSetup {
		private final Poi poi;

		private PoiOperationSetup(Poi poi) {
			this.poi = poi;
		}

		public PoiCellReader readAt(PoiCellRef cellRef) {
			return new PoiCellReader(poi.sheet, cellRef.cellReference());
		}

		public PoiColumnReader readColumns(PoiAreaRef areaRef) {
			return new PoiColumnReader(poi.sheet, areaRef.columnReference());
		}

		public PoiColumnReader readColumns(PoiAreaRef areaRef, int columnIndex) {
			return new PoiColumnReader(poi.sheet, areaRef.columnReference(columnIndex));
		}

		public PoiAreaReader readAll(PoiAreaRef areaRef) {
			return new PoiAreaReader(poi.sheet, areaRef.areaReference());
		}

		public PoiCellWriter writeAt(PoiCellRef cellRef) {
			return new PoiCellWriter(poi.sheet, cellRef.cellReference());
		}

		public PoiColumnWriter writeColumns(PoiAreaRef areaRef) {
			return new PoiColumnWriter(poi.sheet, areaRef.columnReference());
		}

		public PoiColumnWriter writeColumns(PoiAreaRef areaRef, int columnIndex) {
			return new PoiColumnWriter(poi.sheet, areaRef.columnReference(columnIndex));
		}

		public PoiAreaWriter writeAll(PoiAreaRef areaRef) {
			return new PoiAreaWriter(poi.sheet, areaRef.areaReference());
		}
	}
}
