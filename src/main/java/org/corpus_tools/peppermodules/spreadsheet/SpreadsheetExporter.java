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

import org.corpus_tools.pepper.common.PepperConfiguration;
import org.corpus_tools.pepper.impl.PepperExporterImpl;
import org.corpus_tools.pepper.modules.PepperExporter;
import org.corpus_tools.pepper.modules.PepperMapper;
import org.corpus_tools.salt.common.SDocument;
import org.corpus_tools.salt.graph.Identifier;
import org.eclipse.emf.common.util.URI;
import org.osgi.service.component.annotations.Component;

@Component(name="SpreadsheetExporterComponent", factory="PepperExporterComponentFactory")
public class SpreadsheetExporter extends PepperExporterImpl implements PepperExporter {
	
	private static final String NAME = "SpreadsheetExporter";

	public SpreadsheetExporter() {
		super();
		setSupplierContact(URI.createFileURI(PepperConfiguration.EMAIL));
		setName(NAME);
		addSupportedFormat("xlsx", "2007+", null);
		setIsMultithreaded(true);
		setProperties(new SpreadsheetExporterProperties());
		setDesc("This exporter transforms a Salt model into a spreadsheet.");
		setDocumentEnding("xlsx");
		setExportMode(EXPORT_MODE.DOCUMENTS_IN_FILES);
	}
	
	@Override
	public PepperMapper createPepperMapper(Identifier identifier) {
		Salt2SpreadsheetMapper mapper = new Salt2SpreadsheetMapper();
		if (identifier.getIdentifiableElement() != null && identifier.getIdentifiableElement() instanceof SDocument) {
			URI resource = getIdentifier2ResourceTable().get(identifier);
			mapper.setResourceURI(resource);
		}
		return mapper;
	}
}
