package org.subtlelib.poiooxml.api.filter;

import java.util.logging.Filter;

/**
 * Created by beni on 10.12.14.
 */
public interface SupportsFilterRendering<T> {
	T setFilterDataRange(FilterDataRange data);
	T filter();
}
