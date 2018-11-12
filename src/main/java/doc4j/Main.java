/**
 * Project:Demo
 * File:Main.java
 * Copyright 2004-2018 Homolo Co., Ltd. All rights reserved.
 */
package doc4j;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class Main {
	
	public static FileService fileService = new FileService();
	
	public static void main(String[] args){
		try {
			List<byte[]> pictures = new ArrayList<byte[]>();
			
			InputStream is1 = new FileInputStream("D:/workspace/maven-workspace/doc4j/src/main/resources/picture1.jpg");
			byte[] picture1  = new byte[is1.available()];
			is1.read(picture1);
	 		is1.close();
	 		pictures.add(picture1);
	 		
			InputStream is2 = new FileInputStream("D:/workspace/maven-workspace/doc4j/src/main/resources/picture2.jpg");
			byte[] picture2  = new byte[is2.available()];
			is2.read(picture2);
	 		is2.close();
	 		pictures.add(picture2);
	 		
	 		FileParam fileParam = new FileParam();
	 		fileParam.setPictures(pictures);
	 		
	 		WordprocessingMLPackage  wmlPackage = createWmlPackage(fileParam);
	 		wmlPackage.save(new File("生成文件.docx"));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static WordprocessingMLPackage createWmlPackage(FileParam fileParam) throws Exception{
		//第1步：加载模板文件
		WmlTool wmlTool = new WmlTool("fileModel1.docx");
		
		//第2步：更新封面信息
		fileService.updateCover(wmlTool, fileParam);
		
		//第3步：生成目录信息
		wmlTool.createCatalog(2);
		
		//第4步：写入摘要信息
		fileService.createAbstract(wmlTool, fileParam);
		
		//第5步：创建图片表格和多级标题
		fileService.createPictures(wmlTool, fileParam);
		
		//第6步：写入结论部分
		fileService.createConclusion(wmlTool, fileParam);
		
		return wmlTool.getWmlPackage();
	}

}
