/**
 * Project:Demo
 * File:HeadingStyle.java
 * Copyright 2004-2018 Homolo Co., Ltd. All rights reserved.
 */
package doc4j;

/**
 * word 鏂囨。涓爣棰樼殑鏍峰紡锛屽惈缂栧彿鏍峰紡鍜屾枃瀛楁牱寮�
 */
public class HeadingStyle {
	
    /**
     * 鏍囬鐨勭瓑绾э紝鈮�1
     */
	private int grade;
	
    /**
     * 鏍囬鎵�鍦ㄦ钀界殑鏁翠綋鏍峰紡锛屽 1銆乭eading 2绛�
     */
	private String pStyle;
	
	/**
	 * 鏍囬缂栫爜鏍煎紡鐨勬鍒欒〃杈惧紡
	 * 鐩墠浠呮敮鎸� #.#锛屽嵆鏁板瓧涔嬮棿浠� . 鍒嗗壊
	 */
	private String hStyle;

	public int getGrade() {
		return grade;
	}

	public void setGrade(int grade) {
		this.grade = grade;
	}

	public String getPStyle() {
		return pStyle;
	}

	public void setPStyle(String pStyle) {
		this.pStyle = pStyle;
	}

	public String getHStyle() {
		return hStyle;
	}

	public void setHStyle(String hStyle) {
		this.hStyle = hStyle;
	}

	public HeadingStyle(int grade, String pStyle) {
		super();
		this.grade = grade;
		this.pStyle = pStyle;
		this.hStyle = "#.#";
	}

	public HeadingStyle(int grade, String pStyle, String hStyle) {
		super();
		this.grade = grade;
		this.pStyle = pStyle;
		this.hStyle = hStyle;
	}

}

