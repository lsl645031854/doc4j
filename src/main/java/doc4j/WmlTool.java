/**
 * Project:Demo
 * File:WmlTool.java
 * Copyright 2004-2018 Homolo Co., Ltd. All rights reserved.
 */
package doc4j;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.math.BigInteger;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.namespace.QName;

import org.docx4j.UnitsOfMeasurement;
import org.docx4j.XmlUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.DocumentModel;
import org.docx4j.model.structure.PageDimensions;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTMarkupRange;
import org.docx4j.wml.CTShd;
import org.docx4j.wml.Color;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.FldChar;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.STBrType;
import org.docx4j.wml.STFldCharType;
import org.docx4j.wml.STHint;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblBorders;
import org.docx4j.wml.TblPr;
import org.docx4j.wml.TblWidth;
import org.docx4j.wml.Tc;
import org.docx4j.wml.TcPr;
import org.docx4j.wml.Text;


/**
 * word 文档的创建工具
 */
public class WmlTool {
	
	/**
	 * 以下备注仅为方便认识 docx4j 相关类
	 */
	//wml：无线置标语言，用以处理word文档创建
	//P：段落 paragraph 
	//R：行 row
	//Br：换行
	//fldChar：Field Charend
	//TOC：内容标题 Title of Content
	//Tbl：表格Table
	//Tc：Table Cell表单元格
	//Tr：Table Row表格的行
	//*Pr：property，属性、样式设置
	
	
	/**
	 * wml 的文档包，文档操作的主要对象
	 */
    private WordprocessingMLPackage wmlPackage;
    
	/**
	 * wml 的文档工厂，是用以创建相关 element 的工具类
	 */
    private ObjectFactory wmlFactory;
    
	/**
	 * wml 的文档主体部分
	 */
    private MainDocumentPart wmlMainPart;
    
	/**
	 * wml 文档内的所有内容
	 */
    private List<Object> wmlAllContents;
    
    /**
     * 标题相关工具包
     */
    private HeadingTool headingTool;
    
	public WordprocessingMLPackage getWmlPackage() {
		return wmlPackage;
	}

	public void setWmlPackage(WordprocessingMLPackage wmlPackage) {
		this.wmlPackage = wmlPackage;
	}

	public ObjectFactory getWmlFactory() {
		return wmlFactory;
	}

	public void setWmlFactory(ObjectFactory wmlFactory) {
		this.wmlFactory = wmlFactory;
	}

	public MainDocumentPart getWmlMainPart() {
		return wmlMainPart;
	}

	public void setWmlMainPart(MainDocumentPart wmlMainPart) {
		this.wmlMainPart = wmlMainPart;
	}

	public List<Object> getWmlAllContents() {
		return wmlAllContents;
	}

	public void setWmlAllContents(List<Object> wmlAllContents) {
		this.wmlAllContents = wmlAllContents;
	}

	public HeadingTool getHeadingTool() {
		return headingTool;
	}

	public void setHeadingTool(HeadingTool headingTool) {
		this.headingTool = headingTool;
	}
	
	/**
	 * 文档工厂、文档包等变量的初始化
	 * @param fileResource 放在resource路径的模板文件名称，含文件后缀
	 */
	public WmlTool(String fileResource) throws FileNotFoundException, Docx4JException {
		URL url = Main.class.getClassLoader().getResource(fileResource);
//		URL url = WmlT6ool.class.getResource(fileResource);
		this.wmlPackage = WordprocessingMLPackage.load(new FileInputStream(new File(url.getPath())));
        this.wmlMainPart = this.wmlPackage.getMainDocumentPart();
        this.wmlAllContents = this.wmlMainPart.getJaxbElement().getBody().getContent();
        this.wmlFactory = Context.getWmlObjectFactory();
        this.headingTool = new HeadingTool();
	}
	
	/**
	 * 获取当前文档下所有表格
	 */
	public List<Tbl> getAllTables(){
		List<Object> elements = getClazzChildren4Element(this.wmlMainPart, Tbl.class);
		
		List<Tbl> tables = new ArrayList<Tbl>(elements.size());
		for(Object element:elements){
			if(element ==null){
				continue;
			}
			tables.add((Tbl)element);
		}
		
		return tables;
	}
	
	/**
	 * 获取当前表格下所有文本域
	 * @param table 表格
	 */
	public static List<Text> getAllTexts4Table(Tbl table){
		List<Object> elements = getClazzChildren4Element(table, Text.class);
		
		List<Text> texts = new ArrayList<Text>(elements.size());
		for(Object element:elements){
			if(element ==null){
				continue;
			}
			texts.add((Text)element);
		}
		
		return texts;
	}
    
	/**
	 * 获取元素下的所有clazz类型子元素，可包含其本身
	 * 例（1）元素 MainDocumentPart 下查找 Tbl.class
	 * （2）元素 Tbl 下查找 Text.class
	 * @param element
	 * @param clazz
	 * @return
	 */
    public static List<Object> getClazzChildren4Element(Object element, Class<?> clazz) {
        List<Object> elements = new ArrayList<Object>();
        if (element instanceof JAXBElement)
            element = ((JAXBElement<?>) element).getValue();

        if(element.getClass().equals(clazz)){
        	elements.add(element);
        }else if (element instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) element).getContent();
            for (Object child : children) {
                elements.addAll(getClazzChildren4Element(child, clazz));
            }
        }
        return elements;
    }
    
    /**
     * 增加分页，从当前行直接跳转到下页
     */
    public void addPageBreak() {
    	addPageBreak(this.wmlMainPart, this.wmlAllContents, this.wmlFactory);
    }
    
    /**
     * 增加分页，从当前行直接跳转到下页
     * @param mainPart 文档主体
     * @param contents 文档内容
     */
    public static void addPageBreak(MainDocumentPart mainPart, List<Object> contents, ObjectFactory factory) {
        Br br = new Br();//换行
        br.setType(STBrType.PAGE);//换页方式
        
        P paragraph = factory.createP();//段落
        paragraph.getContent().add(br);
        
        contents.add(paragraph);
    }
    
    /**
     * 复杂字符域上边界
     * @param paragraph
     */
    public void addFieldBegin(P paragraph) {
    	addFieldBegin(paragraph, this.wmlFactory);
    }
    
    /**
     * 复杂字符域上边界
     * @param paragraph
     * @param factory
     */
    public static void addFieldBegin(P paragraph, ObjectFactory factory) {
        FldChar fldChar = factory.createFldChar();//字符域
        fldChar.setFldCharType(STFldCharType.BEGIN);
        fldChar.setDirty(true);
        
        R row = factory.createR();//行
        row.getContent().add(getWrappedFldChar(fldChar));
        paragraph.getContent().add(row);
    }
    
    /**
     * 复杂字符域下边界
     * @param paragraph
     */
    public void addFieldEnd(P paragraph) {
    	addFieldEnd(paragraph, this.wmlFactory);
    }
    
    /**
     * 复杂字符域下边界
     * @param paragraph
     * @param factory
     */
    public static void addFieldEnd(P paragraph, ObjectFactory factory) {
        FldChar fldChar = factory.createFldChar();//字符域
        fldChar.setFldCharType(STFldCharType.END);
        
        R row = factory.createR();//行
        row.getContent().add(getWrappedFldChar(fldChar));
        paragraph.getContent().add(row);
    }
    
    /**
     * 包装复杂字符域
     * @param fldChar 字符域
     * @return
     */
    public static JAXBElement getWrappedFldChar(FldChar fldChar) {
    	QName qName = new QName(Namespaces.NS_WORD12, "fldChar");
    	JAXBElement element = new JAXBElement(qName, FldChar.class, fldChar);
        return element;
    }
    
    /**
     * 创建目录信息
     * @param showMaxGrade 目录显示的最大等级
     */
    public void createCatalog(int showMaxGrade) throws JAXBException{
    	createCatalog(this.wmlPackage , this.wmlFactory, showMaxGrade);
    }
    
    /**
     * 创建目录信息
     * @param _package 文档包
     * @param factory 文档工具类
     * @param showMaxGrade 目录显示的最大等级
     * @throws JAXBException
     */
    public static void createCatalog(WordprocessingMLPackage _package , ObjectFactory factory, int showMaxGrade) throws JAXBException {
    	MainDocumentPart mainPart = _package.getMainDocumentPart();
    	List<Object> allContents = mainPart.getJaxbElement().getBody().getContent();
    	
    	addPageBreak(mainPart, allContents, factory);
    	
    	StringBuilder xml = new StringBuilder();
    	xml.append("<w:p w:rsidR='00C54076' w:rsidRDefault='00C54076' ");
    	xml.append("	xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'> ");
    	xml.append("    <w:pPr> ");
    	xml.append("        <w:pStyle w:val='TOC'/> ");
    	xml.append("    </w:pPr> ");
    	xml.append("    <w:r> ");
    	xml.append("        <w:t>目录</w:t> ");
    	xml.append("    </w:r> ");
    	xml.append("</w:p> ");

        P toc = (P) XmlUtils.unmarshalString(xml.toString());//目录标题Title of catalog
        _package.getMainDocumentPart().addObject(toc);

        P catalog = factory.createP();
        addFieldBegin(catalog, factory);
        addTable4Catalog(catalog, showMaxGrade, factory);
        addFieldEnd(catalog, factory);
        
        allContents.add(catalog);
    }
    
    /**
     * 添加目录详情，word内alt + f9可见
     * @param catalog 目录段落
     * @param grade 目录显示的最大等级
     * @param factory 文档工具类
     */
    public static void addTable4Catalog(P catalog, int grade, ObjectFactory factory) {
        Text txt = new Text();
        txt.setSpace("preserve");//目录中保留 heading 前后的空格
        txt.setValue("TOC \\o \"1-"+grade+"\" \\h \\z \\u");
        
        R row = factory.createR();
        row.getContent().add(factory.createRInstrText(txt));
        
        catalog.getContent().add(row);
    }
    
    /**
     * 以指定文本样式创建段落标题 Title of Content，并以默认方式生成编码
     * 注意：一级标题应当声明为1，其他标题方可声明Heading*
     * @param grade 标题等级
     * @param headingText 标题的文本信息（不含编号）
     * @param headingStyle 标题的文本信息（不含编号）
     * @throws JAXBException
     */
    public void addHeading(String headingText, HeadingStyle headingStyle) throws JAXBException{
    	addHeading(headingText, headingStyle, new HeadingFormat("#.#")) ;
    }
    
    /**
     * 创建段落标题 Title of Content，当前文档自增编号
     * 注意：一级标题应当声明为1，其他标题方可声明Heading*
     * @param headingFormat 标题编号格式化
     * @param headingText 标题的文本信息（不含编号）
     * @param headingStyle 标题的样式
     * @throws JAXBException
     */
    public void addHeading(String headingText, HeadingStyle headingStyle, HeadingFormat headingFormat ) throws JAXBException{
    	int grade = headingStyle.getGrade();
    	String headingNum = headingTool.getAutoHeadingNum(grade, headingFormat);
    	addHeading(headingNum, headingText, headingStyle);
    	
    	//通过直接调用方法 addHeading(String headingNum, String headingText, HeadingStyle headingStyle) [以下简称方法1]创建标题，不会自动更新 标题组Map
    	//本方法 [以下简称方法2]因为使用 标题组Map 生成编码，故调用一次更新一次 标题组Map
    	//对同一 wmlTool 的操作，如果时而调用方法1，时而调用方法2，这会导致标题等级混乱
    	//TODO 区分开 生成编码方式，再在调用方法1 时更新或创建新 标题组Map，这要使用到 HeadingFormat.parse方法
    }
    
    /**
     * 创建段落标题 Title of Content
     * @param headingNum 标题的编号
     * @param headingText 标题的文本信息（不含编号）
     * @param headingStyle 标题的样式
     */
    private void addHeading(String headingNum, String headingText, HeadingStyle headingStyle) throws JAXBException {
    	String heading = headingNum + " " +headingText;
    	heading = heading.trim();
    	
    	String pStyle = headingStyle.getPStyle();
    	int blank = headingStyle.getGrade();
        while (blank > 1) {// 按级缩进
            heading = " " + heading;
            blank--;
        }
        
        StringBuilder xml = new StringBuilder();
        xml.append("<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ");
        xml.append("    w:rsidRDefault='00C54076' w:rsidP='00C54076' w:rsidR='00C54076'> ");
        xml.append("    <w:pPr> ");
        xml.append("        <w:pStyle w:val='" + pStyle + "'/> ");
        xml.append("        <w:numPr> ");
        xml.append("            <w:ilvl w:val='0'/> ");
        xml.append("            <w:numId w:val='0'/> ");//如果pStyle是Heading类，不采用自带编号
        xml.append("        </w:numPr> ");
        xml.append("        <w:rPr> ");
        xml.append("            <w:rFonts w:eastAsia='黑体' w:ascii='黑体'/> ");
        xml.append("            <w:b w:val='false'/> ");
        xml.append("            <w:sz w:val='21'/> ");
        xml.append("            <w:szCs w:val='21'/> ");
        xml.append("        </w:rPr> ");
        xml.append("    </w:pPr> ");
        xml.append("    <w:bookmarkStart w:name='_Toc343672792' w:id='0'/> ");
        xml.append("    <w:r> ");
        xml.append("        <w:rPr> ");
        xml.append("            <w:rFonts w:eastAsia='黑体' w:ascii='黑体' w:hint='eastAsia'/> ");
        xml.append("            <w:color w:val='000000'/> ");
        xml.append("            <w:b w:val='false'/> ");
        xml.append("            <w:sz w:val='21'/> ");
        xml.append("            <w:szCs w:val='21'/> ");
        xml.append("        </w:rPr> ");
        xml.append("        <w:t xml:space='preserve'>" + heading + "</w:t> ");//属性preserve 保留空格
        xml.append("    </w:r> ");
        xml.append("    <w:bookmarkEnd w:id='0'/> ");
        xml.append("</w:p> ");
        
        P paragraph = (P) XmlUtils.unmarshalString(xml.toString());
        
        this.wmlMainPart.addObject(paragraph);
    }
    
    /**
     * 创建段落标题 Title of Content
     * 建议使用方法 addHeading(String headingNum, String headingText, HeadingStyle headingStyle);
     */
    @Deprecated
    private void createHeading(String headingNum, String headingText, HeadingStyle headingStyle){
    	
    	String rowTextVal = headingNum + " " +headingText;
    	rowTextVal = rowTextVal.trim();
    	
    	int grade = headingStyle.getGrade();
    	int blank = grade;
        while (blank > 1) {// 按级缩进
            rowTextVal = " " + rowTextVal;
            blank--;
        }
        
    	RPr rowPr = setFontPr4Row(this.wmlFactory, "微软雅黑", 21, false, "000000");//黑色
    	
    	R row = this.wmlFactory.createR();//行
		row.setRPr(rowPr);
		
		Text rowText = this.wmlFactory.createText();//行的文本内容
		rowText.setValue(rowTextVal);
		rowText.setSpace("preserve");//保留文本内容的中的空格
		row.getContent().add(rowText);
		
		PStyle pStyle = this.wmlFactory.createPPrBasePStyle();//段落样式
		pStyle.setVal("Heading" + grade);//目录分级
		
		Jc jc = this.wmlFactory.createJc();
		jc.setVal(JcEnumeration.LEFT);//标题居左
		
		PPr pPr = this.wmlFactory.createPPr();//段落属性
		pPr.setPStyle(pStyle);
		pPr.setJc(jc);
		
		//不用自带编码：pPr.setNumPr(PPrBase.NumPr.NumId=0);
		
		P paragraph = this.wmlFactory.createP();//段落
		paragraph.getContent().add(row);
		paragraph.setPPr(pPr);

		this.wmlMainPart.addObject(paragraph);
    }
    
    /**
     * 添加字符段落
     * @param text 文本信息
     */
    public void addTextParagraph(String text){
    	addTextParagraph(text, this.wmlMainPart);
    }
    
    /**
     * 添加字符段落
     * @param text 文本信息
     * @param mainPart 文档主体
     */
    public static void addTextParagraph(String text, MainDocumentPart mainPart){
    	mainPart.addParagraphOfText(text);
    }
    
    /**
     * 在文档主体添加元素
     * @param element
     */
    public void addElement(Object element){
    	this.wmlMainPart.addObject(element);
    }
    
    /**
     * 设置表格边框颜色及文字对齐方式
     */
    public static void addTblBorders(Tbl table) {
        CTBorder border = new CTBorder();
        border.setColor("auto");
        border.setSz(new BigInteger("4"));
        border.setSpace(new BigInteger("0"));
        border.setVal(STBorder.SINGLE);

        TblBorders borders = new TblBorders();
        borders.setTop(border);//上边框
        borders.setBottom(border);//下边框
        borders.setLeft(border);//左边框
        borders.setRight(border);//右边框
        
        borders.setInsideH(border);//纵向内边框
        borders.setInsideV(border);//横向内边框
        
        TblPr tblPr = new TblPr();//表格属性
        tblPr.setTblBorders(borders);
        table.setTblPr(tblPr);

        Jc jc = new Jc();//文字对象方式
        jc.setVal(JcEnumeration.CENTER);
        
        table.getTblPr().setJc(jc);
    }
    
    /**
     * 添加表单元格
     * @param cellText 单元格文本
     * @param width 单元格宽度
     * @return
     */
    public Tc createTblCell(String cellText, Integer width){
    	return createTblCell( cellText, width, this.wmlFactory);
    }
    
    /**
     * 添加表单元格
     * @param cellText 单元格文本
     * @param width 单元格宽度
     * @param factory 文档工具类
     * @return
     */
    public static Tc createTblCell(String cellText, Integer width, ObjectFactory factory) {
        Text text = factory.createText();
        text.setValue(cellText);
        
        R row = factory.createR();
        row.getContent().add(text);
        
        P paragraph = factory.createP();
        paragraph.getContent().add(row);
    	
        Tc tblCell = factory.createTc();
        tblCell.getContent().add(paragraph);

        if (width != null) {
            TblWidth cellWidth = factory.createTblWidth();
            cellWidth.setType("dxa");//横向宽度
            cellWidth.setW(BigInteger.valueOf(width));
            
            TcPr tcPr = factory.createTcPr();//Table cell property
            tcPr.setTcW(cellWidth);
            
            tblCell.setTcPr(tcPr);
        }
        return tblCell;
    }
    
    /**
     * 设置表格宽度
     * @param table 表格
     * @param width 宽度
     */
    public static void setTblWidth(Tbl table, int width) {
        TblPr tblPr = table.getTblPr();
        if (tblPr == null) {
            tblPr = new TblPr();
            table.setTblPr(tblPr);
        }

        TblWidth tblW = tblPr.getTblW();
        if (tblW == null) {
            tblW = new TblWidth();
            tblPr.setTblW(tblW);
        }
        tblW.setType("dxa");//横向宽度
        tblW.setW(new BigInteger(width+""));
    }
    
    /**
     * 设置单元格背景色
     * @param tblCell 单元格
     * @param colorStr 颜色，如：FFFF00
     */
    public void setTblCellColor(Tc tblCell, String colorStr){
    	setTblCellColor(tblCell, colorStr, this.wmlFactory);
    }
    
    /**
     * 设置单元格背景色
     * @param tblCell 单元格
     * @param colorStr 颜色，如：FFFF00
     * @param factory 文档工具类
     */
    public static void setTblCellColor(Tc tblCell, String colorStr, ObjectFactory factory) {
    	CTShd shd = factory.createCTShd();
    	shd.setFill(colorStr);
    	
        TcPr tblCellPro = factory.createTcPr();
        tblCellPro.setShd(shd);
        tblCell.setTcPr(tblCellPro);
    }
    
    /**
     * 获取文档的可用宽度
     */
    public int getWritableWidth() throws NullPointerException{
    	return getWritableWidth(this.wmlPackage);
    }
    
    /**
     * 获取文档的可用宽度
     * @param _package 文档包
     */
    public static int getWritableWidth(WordprocessingMLPackage _package) throws NullPointerException {
    	DocumentModel docModel = _package.getDocumentModel();//模板文件
    	List<SectionWrapper> sectionWrappers = docModel.getSections();//包装器
    	if(sectionWrappers !=null && sectionWrappers.size()>0){
    		PageDimensions pageDim = sectionWrappers.get(0).getPageDimensions();//页面尺寸
    		return pageDim.getWritableWidthTwips();
    	}
    	
        throw new NullPointerException();
    }
    
    /**
     * 添加书签
     * @param id 书签id
     * @param name 书签名称
     * @param paragraph 段落
     * @param row 行
     * @param factory 文档工具类
     */
    public void addBookMark(String name,P paragraph, R row){
    	addBookMark( 0, name, paragraph, row, this.wmlFactory);
    }
    
    /**
     * 添加书签
     * @param id 书签id
     * @param name 书签名称
     * @param paragraph 段落
     * @param row 行
     * @param factory 文档工具类
     */
    public static void addBookMark( int id, String name,P paragraph, R row, ObjectFactory factory) throws ArrayIndexOutOfBoundsException{
        int index = paragraph.getContent().indexOf(row);

        if (index < 0) {
            throw new ArrayIndexOutOfBoundsException("The current Row does not exist！");
        }

        BigInteger _id = BigInteger.valueOf(id);

        CTMarkupRange mr = factory.createCTMarkupRange();
        mr.setId(_id);
        JAXBElement<CTMarkupRange> bmEnd = factory.createBodyBookmarkEnd(mr);
        paragraph.getContent().add(index + 1, bmEnd);

        CTBookmark bm = factory.createCTBookmark();
        bm.setId(_id);
        bm.setName(name);
        JAXBElement<CTBookmark> bmStart = factory.createBodyBookmarkStart(bm);
        paragraph.getContent().add(index, bmStart);
    }
    
    /**
     * 创建含有内联图片的段落
     * @param picBytes 图片文件流
     * @param width 图片宽度
     * @param height 图片高度
     */
    public P createPWithPicture(byte[] picBytes, long width, long height) throws Exception{
    	return createPWithPicture(picBytes, width, height, this.wmlPackage, this.wmlFactory);
    }
    
    /**
     * 创建含有内联图片的段落
     * @param picBytes 图片文件流
     * @param width 图片宽度
     * @param height 图片高度
     * @param _package 文档包
     * @param factory 文档工具类
     */
    public static P createPWithPicture(byte[] picBytes, long width, long height,
    		WordprocessingMLPackage _package, ObjectFactory factory) throws Exception{
    	
    	BinaryPartAbstractImage picPart = BinaryPartAbstractImage.createImagePart(_package, picBytes);
        Inline inline = picPart.createImageInline("", "", 1, 2, UnitsOfMeasurement.twipToEMU(width),
                UnitsOfMeasurement.twipToEMU(height), false);

        Drawing drawing = factory.createDrawing();
        drawing.getAnchorOrInline().add(inline);

        R row = factory.createR();
        row.getContent().add(drawing);

        P paragraph = factory.createP();
        paragraph.getContent().add(row);

        return paragraph;
    }
    
    /**
     * 设置行文字字体、大小、加粗、颜色
     * @param factory wml 的文档工厂，是用以创建相关 element 的工具类
     * @param fontVal 字体名称，如黑体、微软雅黑、宋体
     * @param fontSize 字体大小，镑，如12
     * @param isBlod 是否加粗
     * @param colorVal 文字颜色，如 FFFF00
     * @return
     */
    private static RPr setFontPr4Row(ObjectFactory factory, String fontVal ,  int fontSize , boolean isBlod , String colorVal){  
        RPr rowPr = factory.createRPr();  
          
        RFonts rowFont = factory.createRFonts();
        rowFont.setHint(STHint.EAST_ASIA);
        rowFont.setAscii(fontVal);
        rowFont.setHAnsi(fontVal);
        rowPr.setRFonts(rowFont);
        
        HpsMeasure _fontSize = factory.createHpsMeasure();
        _fontSize.setVal(BigInteger.valueOf(fontSize));
        rowPr.setSz(_fontSize);
        rowPr.setSzCs(_fontSize);
          
        BooleanDefaultTrue fontBold = factory.createBooleanDefaultTrue();  
        rowPr.setBCs(fontBold);  
        if(isBlod){  
        	rowPr.setB(fontBold);  
        }
        
        Color color = factory.createColor();
        color.setVal(colorVal);
        rowPr.setColor(color);
        
        return rowPr;  
    } 
    
}


