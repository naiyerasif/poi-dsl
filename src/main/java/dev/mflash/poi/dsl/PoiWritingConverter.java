package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Row;

@FunctionalInterface
public interface PoiWritingConverter<S> {

	void convert(S value, Row row);
}
