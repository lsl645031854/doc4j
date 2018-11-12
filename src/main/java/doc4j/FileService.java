/**
 * Project:Demo
 * File:FileService.java
 * Copyright 2004-2018 Homolo Co., Ltd. All rights reserved.
 */
package doc4j;

import java.util.List;

import javax.xml.bind.JAXBException;

import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

/**
 * 文件相关的Service类
 */
public class FileService {
	
    /**
     * 第2步：更新封面
     * @param wmlTool
     * @param fileParam
     */
    public void updateCover(WmlTool wmlTool, FileParam fileParam) {
        List<Tbl> docTbls = wmlTool.getAllTables();
        Tbl table1 = docTbls.get(0);
        
        List<Text> tbl1Txts = WmlTool.getAllTexts4Table(table1);
        Text tbl1Txt1 = tbl1Txts.get(0);
        
        String tbl1Txt1Key = tbl1Txt1.getValue();
        if (tbl1Txt1Key.equals("key0")) {
        	tbl1Txt1.setValue("文档标题（表1行1）");
        }

        Tbl table2 = docTbls.get(1);
        List<Text> tbl2Txts = WmlTool.getAllTexts4Table(table2);
        for(Text tbl2Txt:tbl2Txts){
        	String tbl2TxtVal = tbl2Txt.getValue();
        	if(tbl2TxtVal.indexOf("key") >=0){
        		tbl2Txt.setValue(tbl2TxtVal.replace("key", "value"));
        	}
        }

    }

    /**
     * 第4步：写入摘要信息
     * @param wmlTool
     * @param fileParam
     */
    public void createAbstract(WmlTool wmlTool, FileParam fileParam) throws JAXBException{
    	wmlTool.addPageBreak();
    	
        wmlTool.addHeading("摘要", new HeadingStyle(1, "1"));

        wmlTool.addTextParagraph("文本内容第1行");
        wmlTool.addTextParagraph("文本内容第2行");

        int tblNIndex =wmlTool.getAllTables().size() - 1;//size - 2 +1，2个封面tbl，1个当前tbl
        String tblNTitle = "表" + tblNIndex + " 我是表标题";
        wmlTool.addTextParagraph(tblNTitle);

        ObjectFactory factory = wmlTool.getWmlFactory();
        
        Tbl tblN = factory.createTbl();
        WmlTool.addTblBorders(tblN);
        wmlTool.addElement(tblN);

        Tr tblRow1 = factory.createTr();
        tblRow1.getContent().add(wmlTool.createTblCell("标题1", 5300));
        tblRow1.getContent().add(wmlTool.createTblCell("标题2", 5300));
        tblN.getContent().add(tblRow1);

        Tr tblRowN = null;
        int rowNum = 0;//行号
        int colNum = 0;//列号
        for (int i=1; i<=7; i++) {
        	colNum = (i%2 == 1 ? 1 : 2);
            if (colNum == 1) {
                tblRowN = factory.createTr();
                rowNum ++;
            }

            tblRowN.getContent().add(wmlTool.createTblCell("我是行" + rowNum + "列" + colNum + "的内容", 5300));
            
            if (colNum ==2 || i ==7) {
                tblN.getContent().add(tblRowN);
            }
        }
        
        wmlTool.addTextParagraph("这是一行");
        wmlTool.addTextParagraph("这又是一行");
        wmlTool.addTextParagraph("可以再来一行");
    }
    
    /**
     * 第5步：创建图片表格和多级标题
     * @param wmlTool
     * @param fileParam
     * @throws Exception
     */
    public void createPictures(WmlTool wmlTool, FileParam fileParam) throws Exception{
    	wmlTool.addPageBreak();
    	
        wmlTool.addHeading("我是1级标题", new HeadingStyle(1, "1"));
        wmlTool.addHeading("我是2级标题", new HeadingStyle(2, "Heading2"));
        
        ObjectFactory factory = wmlTool.getWmlFactory();
        Tbl picTbl = factory.createTbl();
        WmlTool.addTblBorders(picTbl);
        WmlTool.setTblWidth(picTbl, 9000);
        
        List<byte[]> pictures = fileParam.getPictures();//至少2张图片
        for (int i=0; i<2; i++) {
            Jc jc = factory.createJc();
            jc.setVal(JcEnumeration.CENTER);
            
            PPr paragraphProperty = factory.createPPr();
            paragraphProperty.setJc(jc);
            
            P paragraph = wmlTool.createPWithPicture(pictures.get(i), 3200, 1800);
            paragraph.setPPr(paragraphProperty);

            Tc tblCell = factory.createTc();
            tblCell.getContent().add(paragraph);

            Tr tblRow = factory.createTr();
            tblRow.getContent().add(tblCell);

            picTbl.getContent().add(tblRow);
        }
        wmlTool.addElement(picTbl);
        
        wmlTool.addHeading("我是2级标题", new HeadingStyle(2, "Heading2"));
    }

    /**
     * 第6步：写入结论部分
     * @param wmlTool
     * @param fileParam
     * @throws JAXBException
     */
    public void createConclusion(WmlTool wmlTool, FileParam fileParam) throws JAXBException{
    	wmlTool.addPageBreak();
        wmlTool.addHeading("结论", new HeadingStyle(1, "1"));

        wmlTool.addTextParagraph("我是结论的详细信息");
    }
    
}
