package me.test.docx4j;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.junit.Test;

public class MyDocx4jTest {

	@Test//是生成文件
	public void test1() {
		try {
			WordprocessingMLPackage pkg =  WordprocessingMLPackage.createPackage();
			MainDocumentPart documentPart = pkg.getMainDocumentPart(); 
			documentPart.addParagraphOfText("Hello Word!");  
			documentPart.addStyledParagraphOfText("Title", "Hello Word!");
			documentPart.addStyledParagraphOfText("Subtitle"," a subtitle!"); 
			pkg.save(new java.io.File("D://xxx.docx")); 
			//pkg.save(new File(System.getProperty("D://游戏"), "Empty.docx"));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
