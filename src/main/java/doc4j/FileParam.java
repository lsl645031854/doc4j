/**
 * Project:Demo
 * File:FileParam.java
 * Copyright 2004-2018 Homolo Co., Ltd. All rights reserved.
 */
package doc4j;

import java.util.List;

/**
 * 与文件的业务相关的参数类
 */
public class FileParam {
	
	private List<byte[]> pictures;

	public List<byte[]> getPictures() {
		return pictures;
	}

	public void setPictures(List<byte[]> pictures) {
		this.pictures = pictures;
	}

}

