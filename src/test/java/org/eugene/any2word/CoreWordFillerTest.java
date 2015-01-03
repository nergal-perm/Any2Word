package org.eugene.any2word;

import static org.junit.Assert.*;

import java.io.File;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Before;
import org.junit.Test;

public class CoreWordFillerTest {
	private WordprocessingMLPackage wordDoc;
	
	@Before
	public void initialize() {
		String resName = "SampleTemplate.docx";
		ClassLoader classLoader = this.getClass().getClassLoader();
		File docxFile = new File(classLoader.getResource(resName).getFile());
		try {
			wordDoc = WordprocessingMLPackage.load(docxFile);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace(System.out);
		}
	}
	
	@Test
	public void testCoreWordFiller() {
		CoreWordFiller target = new CoreWordFiller(wordDoc);
		assertNotNull(target.getWordDocument());
		assertNotNull(target.getWordDocument().getMainDocumentPart());
	}

}
