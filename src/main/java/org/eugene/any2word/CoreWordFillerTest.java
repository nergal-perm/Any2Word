package org.eugene.any2word;

import static org.junit.Assert.*;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.Test;

public class CoreWordFillerTest {
	private CoreWordFiller target;
	

	@Test
	public void testLoadFileUponCreation() {
		target = new CoreWordFiller(new WordprocessingMLPackage());
		assertNotNull(target.getWordDocument());
	}

}
