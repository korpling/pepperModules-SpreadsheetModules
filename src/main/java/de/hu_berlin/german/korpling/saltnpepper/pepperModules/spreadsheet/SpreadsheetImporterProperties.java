package de.hu_berlin.german.korpling.saltnpepper.pepperModules.spreadsheet;

import org.corpus_tools.pepper.modules.PepperModuleProperties;
import org.corpus_tools.pepper.modules.PepperModuleProperty;

@SuppressWarnings("serial")
public class SpreadsheetImporterProperties extends PepperModuleProperties {

	public static final String PROP_CONCATENATE_TEXT = "concatenateText";

	public SpreadsheetImporterProperties() {
		addProperty(new PepperModuleProperty<>(PROP_CONCATENATE_TEXT, Boolean.class, "Defines, if the textual data shall be concatenated or if a new string object shall be created (key: 'concatenateText', default is 'true').", true, false));
		}

	public Boolean concatenateText() {
		return (Boolean) getProperty(PROP_CONCATENATE_TEXT).getValue();
	}
}
