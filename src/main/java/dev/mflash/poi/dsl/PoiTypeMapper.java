package dev.mflash.poi.dsl;

import org.apache.poi.ss.usermodel.Row;

@FunctionalInterface
public interface PoiTypeMapper<R extends Row, S> {

	Row mapType(S value);
}
