/**
 * Copyright 2009 Humboldt-Universität zu Berlin, INRIA.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 *
 */
package org.corpus_tools.peppermodules.spreadsheet;

import org.corpus_tools.pepper.modules.PepperModuleProperties;
import org.corpus_tools.pepper.modules.PepperModuleProperty;


public class SpreadsheetImporterProperties extends PepperModuleProperties {

	/**
	 * 
	 */
	private static final long serialVersionUID = 6193481351811254201L;
	public static final String PROP_PRIMARY_TEXT = "primText";
	public static final String PROP_CORPUS_SHEET = "corpusSheet";
	public static final String PROP_META_SHEET = "metaSheet";
	public static final String PROP_ANNO_REFERS_TO = "annoPrimRel";
	public static final String PROP_SET_LAYER = "setLayer";
	public static final String PROP_META_ANNO = "metaAnnotation";
	public static final String PROP_INCLUDE_EMPTY_PRIM_CELLS = "includeEmptyPrimCells";
	public static final String PROP_ADD_ORDER_RELATION = "addOrderRelation";
	public static final String PROP_ANNO_SHORT_PRIM_REL = "shortAnnoPrimRel";
//	public static final String PROP_USE_ANNO_FOR_PRIM_DETECTION = "useAnnoForPrimDetection";

	public SpreadsheetImporterProperties() {
		addProperty(new PepperModuleProperty<>(PROP_PRIMARY_TEXT, String.class, "Defines the name of the column(s), that hold the primary text, this can either be a single column name, or a comma seperated enumeration of column names (key: 'primText', default value: 'tok').", "tok", false));
		addProperty(new PepperModuleProperty<>(PROP_CORPUS_SHEET, String.class, "Defines the sheet, that holds the actual corpus information (key: 'corpusSheet', default is 'Tabelle1'). By default the first sheet of your spreadsheets will be used (independent of the name).", "Tabelle1", false));
		addProperty(new PepperModuleProperty<>(PROP_ANNO_REFERS_TO, String.class, "Defines which annotation tier refers to which primary text tier in a direct way, therefor the annotation tier name is followed by the name of the tier, that holds the primary text in square brackets. A possible key-value set could be: key='annoPrimRel', value='anno1=anno1[tok1]' (key: 'annoPrimRel', default is 'null').", null, false));
		addProperty(new PepperModuleProperty<>(PROP_SET_LAYER, String.class, "Defines the corresponding layer to the given annotation layer (key: 'setLayer') ", null, false));
		addProperty(new PepperModuleProperty<>(PROP_META_SHEET, String.class, "Defines the sheet, that holds the meta information of the document (key: 'metaSheet', default is 'Tabelle2'). By default the second sheet of your spreadsheets will be used (independent of the name).", "Tabelle2", false));
		addProperty(new PepperModuleProperty<>(PROP_META_ANNO, Boolean.class, "Defines, if another sheet shall be scaned for meta annotation (key: 'metaAnnotation', default is 'true').", true, false));
		addProperty(new PepperModuleProperty<>(PROP_INCLUDE_EMPTY_PRIM_CELLS, Boolean.class, "If true allow empty cells of primary text, default is 'false'.",  false, false));
		addProperty(new PepperModuleProperty<>(PROP_ADD_ORDER_RELATION, Boolean.class, "If true add an explicit order relation between each primary text token. Default is 'true'.",  true, false));
		addProperty(new PepperModuleProperty<>(PROP_ANNO_SHORT_PRIM_REL, String.class, "Defines which primary text tiers are the basis of which annotation tiers, therefor a comma seperated list of primary text tiers, followed by a list of all annotations that refer to the primary tier, is needed. A possible key-value set could be: key='shortAnnoPrimRel', value='primText1={anno1, anno2}, primText2={anno3}' (key: 'annoPrimRel', default is 'null').",  null, false));
//		addProperty(new PepperModuleProperty<>(PROP_USE_ANNO_FOR_PRIM_DETECTION, Boolean.class, "If true use the annotations to detect the related primary text of an annotation. Therefore the annotation should have the following form: annoName[relatedPrimaryText]. Default is 'true'.",  true, false));
	}

	public String getPrimaryText() {
		return (String) getProperty(PROP_PRIMARY_TEXT).getValue();
	}

	public String getCorpusSheet() {
		return (String) getProperty(PROP_CORPUS_SHEET).getValue();
	}
	
	public String getMetaSheet() {
		return (String) getProperty(PROP_META_SHEET).getValue();
	}
	
	public String getAnnoPrimRel() {
		return (String) getProperty(PROP_ANNO_REFERS_TO).getValue();
	}
	
	public String getShortAnnoPrimRel() {
		return (String) getProperty(PROP_ANNO_SHORT_PRIM_REL).getValue();
	}
	
	public String getLayer(){
		return (String) getProperty(PROP_SET_LAYER).getValue();
	}
	
	public Boolean getMetaAnnotation() {
		return (Boolean) getProperty(PROP_META_ANNO).getValue();
	}
	
	public Boolean getIncludeEmptyPrimCells() {
		return (Boolean) getProperty(PROP_INCLUDE_EMPTY_PRIM_CELLS).getValue();
	}
	
	public Boolean getAddOrderRelation() {
		return (Boolean) getProperty(PROP_ADD_ORDER_RELATION).getValue();
	}
	
//	public Boolean getUseAnnoForPrimDetection() {
//		return (Boolean) getProperty(PROP_USE_ANNO_FOR_PRIM_DETECTION).getValue();
//	}
}
