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
import org.corpus_tools.pepper.impl.PepperImporterImpl;
import org.corpus_tools.pepper.modules.PepperImporter;
import org.corpus_tools.pepper.modules.PepperMapper;
import org.corpus_tools.pepper.modules.PepperModule;
import org.corpus_tools.pepper.modules.PepperModuleProperties;
import org.corpus_tools.pepper.modules.exceptions.PepperModuleNotReadyException;
import org.corpus_tools.salt.common.SCorpus;
import org.corpus_tools.salt.common.SDocument;
import org.corpus_tools.salt.graph.Identifier;
import org.eclipse.emf.common.util.URI;
import org.osgi.service.component.annotations.Component;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Vivian Voigt
 * @version 1.0
 *
 */
@Component(name="SpreadsheetImporterComponent", factory="PepperImporterComponentFactory")
public class SpreadsheetImporter extends PepperImporterImpl implements PepperImporter{
// =================================================== mandatory ===================================================
	public static final String MODULE_NAME="SpreadsheetImporter";
	
	// this is a logger, for recording messages during program process, like debug messages
	static final Logger logger= LoggerFactory.getLogger(MODULE_NAME);
	
	/**
	 * <strong>OVERRIDE THIS METHOD FOR CUSTOMIZATION</strong>
	 * <br/>
	 * A constructor for your module. Set the coordinates, with which your module shall be registered. 
	 * The coordinates (modules name, version and supported formats) are a kind of a fingerprint, 
	 * which should make your module unique.
	 */
	public SpreadsheetImporter(){
		super();
		setSupplierContact(URI.createFileURI(PepperConfiguration.EMAIL));
		setName(MODULE_NAME);
		addSupportedFormat("xls", "97-2008", null);
		addSupportedFormat("xlsx", "2007+", null);
		getDocumentEndings().add("xlsx");
		getDocumentEndings().add("xls");
		this.setProperties(new SpreadsheetImporterProperties());
	}
	
	
	/**
	 * <strong>OVERRIDE THIS METHOD FOR CUSTOMIZATION</strong>
	 * <br/>
	 * This method creates a customized {@link PepperMapper} object and returns it. You can here do some additional initialisations. 
	 * Thinks like setting the {@link Identifier} of the {@link SDocument} or {@link SCorpus} object and the {@link URI} resource is done
	 * by the framework (or more in detail in method {@link #start()}).<br/> 
	 * The parameter <code>Identifier</code>, if a {@link PepperMapper} object should be created in case of the object to map is either 
	 * an {@link SDocument} object or an {@link SCorpus} object of the mapper should be initialized differently. 
	 * <br/>
	 * Just to show how the creation of such a mapper works, we here create a sample mapper of type {@link Spreadsheet2SaltMapper}, 
	 * which only produces a fixed document-structure in method  {@link Spreadsheet2SaltMapper#mapSDocument()} and enhances the 
	 * corpora for further meta-annotations in the method {@link Spreadsheet2SaltMapper#mapSCorpus()}.
	 * <br/>
	 * If your mapper needs to have set variables, this is the place to do it.
	 * @param Identifier {@link Identifier} of the {@link SCorpus} or {@link SDocument} to be processed. 
	 * @return {@link PepperMapper} object to do the mapping task for object connected to given {@link Identifier}
	 */
	@Override
	public PepperMapper createPepperMapper(Identifier identifier){
		Spreadsheet2SaltMapper mapper= new Spreadsheet2SaltMapper();
		if (identifier.getIdentifiableElement() != null && identifier.getIdentifiableElement() instanceof SDocument) {
			URI resource = getIdentifier2ResourceTable().get(identifier);
			mapper.setResourceURI(resource);
		}
		return mapper;
	}
	
	/**
	 * <strong>OVERRIDE THIS METHOD FOR CUSTOMIZATION</strong>
	 * <br/>
	 * This method is called by the pepper framework and returns if a corpus located at the given {@link URI} is importable
	 * by this importer. If yes, 1 must be returned, if no 0 must be returned. If it is not quite sure, if the given corpus
	 * is importable by this importer any value between 0 and 1 can be returned. If this method is not overridden, 
	 * null is returned.
	 * @return 1 if corpus is importable, 0 if corpus is not importable, 0 < X < 1, if no definitive answer is possible,  null if method is not overridden 
	 */
	

// =================================================== optional ===================================================	
	/**
	 * <strong>OVERRIDE THIS METHOD FOR CUSTOMIZATION</strong>
	 * <br/>
	 * This method is called by the pepper framework after initializing this object and directly before start processing. 
	 * Initializing means setting properties {@link PepperModuleProperties}, setting temporary files, resources etc.
	 * returns false or throws an exception in case of {@link PepperModule} instance is not ready for any reason.
	 * <br/>
	 * So if there is anything to do, before your importer can start working, do it here.
	 * @return false, {@link PepperModule} instance is not ready for any reason, true, else.
	 */
	@Override
	public boolean isReadyToStart() throws PepperModuleNotReadyException{
		//TODO make some initializations if necessary
		return(super.isReadyToStart());
	}

}
