package org.subtlelib.poiooxml.api.configuration;

/**
 * Configuration provider.
 * 
 * @author i.voshkulat
 *
 */
public interface ConfigurationProvider {

	/**
	 * Get workbook wide configuration.
	 * 
	 * @return configuration
	 */
	public Configuration getConfiguration();

}
