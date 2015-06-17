package de.hu_berlin.german.korpling.saltnpepper.pepperModules.spreadsheet;

import org.eclipse.emf.common.util.BasicEList;
import org.eclipse.emf.common.util.EList;
import org.eclipse.emf.common.util.URI;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import de.hu_berlin.german.korpling.saltnpepper.pepper.common.DOCUMENT_STATUS;
import de.hu_berlin.german.korpling.saltnpepper.pepper.modules.impl.PepperMapperImpl;
import de.hu_berlin.german.korpling.saltnpepper.salt.SaltFactory;
import de.hu_berlin.german.korpling.saltnpepper.salt.saltCommon.sDocumentStructure.SPointingRelation;
import de.hu_berlin.german.korpling.saltnpepper.salt.saltCommon.sDocumentStructure.SSpan;
import de.hu_berlin.german.korpling.saltnpepper.salt.saltCommon.sDocumentStructure.SStructure;
import de.hu_berlin.german.korpling.saltnpepper.salt.saltCommon.sDocumentStructure.STYPE_NAME;
import de.hu_berlin.german.korpling.saltnpepper.salt.saltCommon.sDocumentStructure.STextualDS;
import de.hu_berlin.german.korpling.saltnpepper.salt.saltCommon.sDocumentStructure.SToken;

/**
 * This class is a dummy implementation for a mapper, to show how it works. This  sample mapper 
 * only produces a fixed document-structure in method  {@link ExcelMapper#mapSDocument()} and enhances the 
 * corpora for further meta-annotations in the method {@link ExcelMapper#mapSCorpus()}.
 * <br/>
 * In production, it might be better to implement the mapper in its own file, we just did it here for 
 * compactness of the code. 
 * @author Florian Zipser
 *
 */
public class ExcelMapper extends PepperMapperImpl{
	/**
	 * 
	 */
	private static final Logger logger= LoggerFactory.getLogger(ExcelMapper.class);
	
	private final ExcelImporter importer;
	/**
	 * @param excelImporter
	 */
	ExcelMapper(ExcelImporter excelImporter) {
		importer = excelImporter;
	}
	/**
	 * <strong>OVERRIDE THIS METHOD FOR CUSTOMIZATION</strong>
	 * <br/>
	 * If you need to make any adaptations to the corpora like adding further meta-annotation, do it here.
	 * When whatever you have done successful, return the status {@link DOCUMENT_STATUS#COMPLETED}. If anything
	 * went wrong return the status {@link DOCUMENT_STATUS#FAILED}.
	 * <br/>
	 * In our dummy implementation, we just add a creation date to each corpus. 
	 */
	@Override
	public DOCUMENT_STATUS mapSCorpus() {
		//getScorpus() returns the current corpus object.
		getSCorpus().createSMetaAnnotation(null, "date", "1989-12-17");
		return(DOCUMENT_STATUS.COMPLETED);
	}
	/**
	 * <strong>OVERRIDE THIS METHOD FOR CUSTOMIZATION</strong>
	 * <br/>
	 * This is the place for the real work. Here you have to do anything necessary, to map a corpus to Salt.
	 * These could be things like: reading a file, mapping the content, closing the file, cleaning up and so on.
	 * <br/>
	 * In our dummy implementation, we do not read a file, for not making the code too complex. We just show how
	 * to create a simple document-structure in Salt, in following steps:
	 * <ol>
	 * 	<li>creating primary data</li>
	 *  <li>creating tokenization</li>
	 *  <li>creating part-of-speech annotation for tokenization</li>
	 *  <li>creating information structure annotation via spans</li>
	 *  <li>creating anaphoric relation via pointing relation</li>
	 *  <li>creating syntactic annotations</li>
	 * </ol>
	 */
	@Override
	public DOCUMENT_STATUS mapSDocument() {
		//the method getSDocument() returns the current document for creating the document-structure
		getSDocument().setSDocumentGraph(SaltFactory.eINSTANCE.createSDocumentGraph());
		//to get the exact resource, which be processed now, call getResources(), make sure, it was set in createMapper()  
		URI resource= getResourceURI();
		
		//we record, which file currently is imported to the debug stream, in this dummy implementation the resource is null 
		logger.debug("Importing the file {}.", resource);
		
		/**
		 * STEP 1:
		 * we create the primary data and hold a reference on the primary data object
		 */
		STextualDS primaryText= getSDocument().getSDocumentGraph().createSTextualDS("Is this example more complicated than it appears to be?");
		
		//we add a progress to notify the user about the process status (this is very helpful, especially for longer taking processes)
		addProgress(0.16);
		
		/**
		 * STEP 2:
		 * we create a tokenization over the primary data
		 */
		SToken tok_is= getSDocument().getSDocumentGraph().createSToken(primaryText, 0,2);		//Is
		SToken tok_thi= getSDocument().getSDocumentGraph().createSToken(primaryText, 3,7);		//this
		SToken tok_exa= getSDocument().getSDocumentGraph().createSToken(primaryText, 8,15);		//example
		SToken tok_mor= getSDocument().getSDocumentGraph().createSToken(primaryText, 16,20);	//more
		SToken tok_com= getSDocument().getSDocumentGraph().createSToken(primaryText, 21,32);	//complicated
		SToken tok_tha= getSDocument().getSDocumentGraph().createSToken(primaryText, 33,37);	//than
		SToken tok_it= getSDocument().getSDocumentGraph().createSToken(primaryText, 38,40);		//it
		SToken tok_app= getSDocument().getSDocumentGraph().createSToken(primaryText, 41,48);	//appears
		SToken tok_to= getSDocument().getSDocumentGraph().createSToken(primaryText, 49,51);		//to
		SToken tok_be= getSDocument().getSDocumentGraph().createSToken(primaryText, 52,54);		//be
		SToken tok_PUN= getSDocument().getSDocumentGraph().createSToken(primaryText, 54,55);	//?
		
		//we add a progress to notify the user about the process status (this is very helpful, especially for longer taking processes)
		addProgress(0.16);
		
		/**
		 * STEP 3:
		 * we create a part-of-speech and a lemma annotation for tokens
		 */
		//we create part-of-speech annotations
		tok_is.createSAnnotation(null, "pos", "VBZ");
		tok_thi.createSAnnotation(null, "pos", "DT");
		tok_exa.createSAnnotation(null, "pos", "NN");
		tok_mor.createSAnnotation(null, "pos", "RBR");
		tok_com.createSAnnotation(null, "pos", "JJ");
		tok_tha.createSAnnotation(null, "pos", "IN");
		tok_it.createSAnnotation(null, "pos", "PRP");
		tok_app.createSAnnotation(null, "pos", "VBZ");
		tok_to.createSAnnotation(null, "pos", "TO");
		tok_be.createSAnnotation(null, "pos", "VB");
		tok_PUN.createSAnnotation(null, "pos", ".");
		
		//we create lemma annotations
		tok_is.createSAnnotation(null, "lemma", "be");
		tok_thi.createSAnnotation(null, "lemma", "this");
		tok_exa.createSAnnotation(null, "lemma", "example");
		tok_mor.createSAnnotation(null, "lemma", "more");
		tok_com.createSAnnotation(null, "lemma", "complicated");
		tok_tha.createSAnnotation(null, "lemma", "than");
		tok_it.createSAnnotation(null, "lemma", "it");
		tok_app.createSAnnotation(null, "lemma", "appear");
		tok_to.createSAnnotation(null, "lemma", "to");
		tok_be.createSAnnotation(null, "lemma", "be");
		tok_PUN.createSAnnotation(null, "lemma", ".");
		
		//we add a progress to notify the user about the process status (this is very helpful, especially for longer taking processes)
		addProgress(0.16);
		
		/**
		 * STEP 4:
		 * we create some information structure annotations via spans, spans can be used, to group tokens to a set, 
		 * which can be annotated
		 * <table border="1">
		 * 	<tr>
		 * 	 	<td>contrast-focus</td>
		 * 		<td colspan="9">topic</td>
		 * </tr>
		 *  <tr>
		 *  	<td>Is</td>
		 *  	<td>this</td>
		 *   	<td>example</td>
		 *    	<td>more</td>
		 *     	<td>complicated</td>
		 *  	<td>than</td>
		 *   	<td>it</td>
		 *    	<td>appears</td>
		 *     	<td>to</td>
		 *     	<td>be</td>
		 *  </tr>
		 * </table>
		 */
		SSpan contrastFocus= getSDocument().getSDocumentGraph().createSSpan(tok_is);
		contrastFocus.createSAnnotation(null, "Inf-Struct", "contrast-focus");
		EList<SToken> topic_set= new BasicEList<SToken>();
		topic_set.add(tok_thi);
		topic_set.add(tok_exa);
		topic_set.add(tok_mor);
		topic_set.add(tok_com);
		topic_set.add(tok_tha);
		topic_set.add(tok_it);
		topic_set.add(tok_app);
		topic_set.add(tok_to);
		topic_set.add(tok_be);
		SSpan topic= getSDocument().getSDocumentGraph().createSSpan(topic_set);
		topic.createSAnnotation(null, "Inf-Struct", "topic");
		
		//we add a progress to notify the user about the process status (this is very helpful, especially for longer taking processes)
		addProgress(0.16);
		
		/**
		 * STEP 5:
		 * we create anaphoric relation between 'it' and 'this example', therefore 'this example' must be added
		 * to a span. This makes use of the graph based model of Salt. First we create a relation, than we
		 * set its source and its target node and last we add the relation to the graph.
		 */
		EList<SToken> target_set= new BasicEList<SToken>();
		target_set.add(tok_thi);
		target_set.add(tok_exa);
		SSpan target= getSDocument().getSDocumentGraph().createSSpan(target_set);
		SPointingRelation anaphoricRel= SaltFactory.eINSTANCE.createSPointingRelation();
		anaphoricRel.setSStructuredSource(tok_is);
		anaphoricRel.setSStructuredTarget(target);
		anaphoricRel.addSType("anaphoric");
		//we add the created relation to the graph
		getSDocument().getSDocumentGraph().addSRelation(anaphoricRel);
		
		//we add a progress to notify the user about the process status (this is very helpful, especially for longer taking processes)
		addProgress(0.16);
		
		/**
		 * STEP 6:
		 * We create a syntax tree following the Tiger scheme
		 */
		SStructure root  = SaltFactory.eINSTANCE.createSStructure();
		SStructure sq    = SaltFactory.eINSTANCE.createSStructure();
		SStructure np1   = SaltFactory.eINSTANCE.createSStructure();
		SStructure adjp1 = SaltFactory.eINSTANCE.createSStructure();
		SStructure adjp2 = SaltFactory.eINSTANCE.createSStructure();
		SStructure sbar  = SaltFactory.eINSTANCE.createSStructure();
		SStructure s1    = SaltFactory.eINSTANCE.createSStructure();
		SStructure np2   = SaltFactory.eINSTANCE.createSStructure();
		SStructure vp1   = SaltFactory.eINSTANCE.createSStructure();
		SStructure s2    = SaltFactory.eINSTANCE.createSStructure();
		SStructure vp2   = SaltFactory.eINSTANCE.createSStructure();
		SStructure vp3   = SaltFactory.eINSTANCE.createSStructure();
		
		//we add annotations to each SStructure node
		root.createSAnnotation(null, "cat", "ROOT");
		sq.createSAnnotation(null, "cat", "SQ");
		np1.createSAnnotation(null, "cat", "NP");
		adjp1.createSAnnotation(null, "cat", "ADJP");
		adjp2.createSAnnotation(null, "cat", "ADJP");
		sbar.createSAnnotation(null, "cat", "SBAR");
		s1.createSAnnotation(null, "cat", "S");
		np2.createSAnnotation(null, "cat", "NP");
		vp1.createSAnnotation(null, "cat", "VP");
		s2.createSAnnotation(null, "cat", "S");
		vp2.createSAnnotation(null, "cat", "VP");
		vp3.createSAnnotation(null, "cat", "VP");
		
		//we add the root node first
		getSDocument().getSDocumentGraph().addSNode(root);
		STYPE_NAME domRel = STYPE_NAME.SDOMINANCE_RELATION;
		//than we add the rest and connect them to each other
		getSDocument().getSDocumentGraph().addSNode(root,  sq,      domRel);
		getSDocument().getSDocumentGraph().addSNode(sq,    tok_is,  domRel); // "Is"
		getSDocument().getSDocumentGraph().addSNode(sq,    np1,     domRel);
		getSDocument().getSDocumentGraph().addSNode(np1,   tok_thi, domRel); // "this"
		getSDocument().getSDocumentGraph().addSNode(np1,   tok_exa, domRel); // "example"
		getSDocument().getSDocumentGraph().addSNode(sq,    adjp1,   domRel);
		getSDocument().getSDocumentGraph().addSNode(adjp1, adjp2,   domRel);
		getSDocument().getSDocumentGraph().addSNode(adjp2, tok_mor, domRel); // "more"
		getSDocument().getSDocumentGraph().addSNode(adjp2, tok_com, domRel); // "complicated"				
		getSDocument().getSDocumentGraph().addSNode(adjp1, sbar,    domRel);
		getSDocument().getSDocumentGraph().addSNode(sbar,  tok_tha, domRel); // "than"
		getSDocument().getSDocumentGraph().addSNode(sbar,  s1,      domRel);				
		getSDocument().getSDocumentGraph().addSNode(s1,    np2,     domRel);				
		getSDocument().getSDocumentGraph().addSNode(np2,   tok_it,  domRel); // "it"
		getSDocument().getSDocumentGraph().addSNode(s1,    vp1,     domRel);
		getSDocument().getSDocumentGraph().addSNode(vp1,   tok_app, domRel); // "appears"				
		getSDocument().getSDocumentGraph().addSNode(vp1,   s2,      domRel);				
		getSDocument().getSDocumentGraph().addSNode(s2,    vp2,     domRel);
		getSDocument().getSDocumentGraph().addSNode(vp2,   tok_to,  domRel); // "to"				
		getSDocument().getSDocumentGraph().addSNode(vp2,   vp3,     domRel);
		getSDocument().getSDocumentGraph().addSNode(vp3,   tok_be,  domRel); // "be"
		getSDocument().getSDocumentGraph().addSNode(root,  tok_PUN, domRel); // "?"
		
		//we set progress to 'done' to notify the user about the process status (this is very helpful, especially for longer taking processes)
		setProgress(1.0);
		
		//now we are done and return the status that everything was successful
		return(DOCUMENT_STATUS.COMPLETED);
	}	
}