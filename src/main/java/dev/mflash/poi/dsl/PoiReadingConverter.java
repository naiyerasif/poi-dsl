package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Row;

@FunctionalInterface
public interface PoiReadingConverter<S> {

	S convert(Row row);
}
