package org.eugene.any2word;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileNotFoundException;

import org.docx4j.XmlUtils;
import org.docx4j.model.datastorage.CustomXmlDataStorage;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Before;
import org.junit.Test;

public class CoreWordFillerTest {
	private WordprocessingMLPackage wordDoc;
	private File xmlFile;

	@Before
	public void initialize() {
		String resDocName = "SampleTemplate.docx";
		String resXmlName = "SampleXML.xml";
		ClassLoader classLoader = this.getClass().getClassLoader();
		File docxFile = new File(classLoader.getResource(resDocName).getFile());
		xmlFile = new File(classLoader.getResource(resXmlName).getFile());
		try {
			wordDoc = WordprocessingMLPackage.load(docxFile);

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace(System.out);
		}
	}

	@Test
	public void testXMLInjectionUponCreation() {
		CoreWordFiller target = null;
		try {
			target = new CoreWordFiller(wordDoc, xmlFile);
		} catch (FileNotFoundException | Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		assertNotNull(target.getWordDocument().getCustomXmlDataStorageParts());
	}

	@Test
	public void testXMLDataCanBeRetrieved() {
		CoreWordFiller target = null;
		String itemId = "";
		try {
			target = new CoreWordFiller(wordDoc, xmlFile);
			itemId = CustomXmlUtils.getCustomXmlItemId(target.getWordDocument()).toLowerCase();
		} catch (Docx4JException | FileNotFoundException e) {
			e.printStackTrace();
		}
		System.out.print(itemId);
		assertNotEquals("Item ID is empty", "", itemId);
	}
}
