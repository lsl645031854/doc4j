/**
 * Project:Demo
 * File:HeadingFormat.java
 * Copyright 2004-2018 Homolo Co., Ltd. All rights reserved.
 */
package doc4j;

import java.util.HashMap;
import java.util.Map;



/**
 * word 文档中标题的编码格式化工具类
 */
public class HeadingFormat {
	
	/**
	 * 编码格式的正则表达式
	 * 目前仅支持 #.#，即数字之间以 . 分割
	 */
	private String pattern;
	
	public HeadingFormat(String pattern) {
		super();
		this.pattern = pattern;
	}

	/**
	 * 给定Map生成等级编码，起始等级为默认1，正则表达式为默认 #.#
	 * @param gradeMaxVals 等级与该等级最大值Map，其中key是等级，value是该等级当前最大值
	 * @param endGrade 显示输出的结束等级
	 * @return
	 */
	public String format(Map<Integer, Integer> gradeMaxVals, int endGrade){
		return format(gradeMaxVals, 1 ,endGrade, this.pattern);
	}
	
	/**
	 * 给定Map生成等级编码，起始等级为默认1
	 * @param gradeMaxVals 等级与该等级最大值Map，其中key是等级，value是该等级当前最大值
	 * @param endGrade 显示输出的结束等级
	 * @param pattern 编码格式的正则表达式
	 * @return
	 */
	public String format(Map<Integer, Integer> gradeMaxVals, int endGrade, String pattern){
		return format(gradeMaxVals, 1 ,endGrade, pattern);
	}
	
	/**
	 * 给定Map生成等级编码
	 * @param gradeMaxVals 等级与该等级最大值Map，其中key是等级，value是该等级当前最大值
	 * @param startGrade 显示输出的起始等级
	 * @param endGrade 显示输出的结束等级
	 * @param pattern 编码格式的正则表达式
	 * @return
	 */
	public String format(Map<Integer, Integer> gradeMaxVals, int startGrade , int endGrade, String pattern){
		if(gradeMaxVals ==null || gradeMaxVals.size() <1){
			return "";
		}
		
		if(startGrade <1){
			throw new ArrayIndexOutOfBoundsException("起始等级应当 ≥1！");
		}
		if(startGrade > endGrade){
			throw new ArrayIndexOutOfBoundsException("起始等级应当小于等于结束等级！");
		}
		
		if(pattern != "#.#"){
			throw new ArrayIndexOutOfBoundsException("非标准格式的暂未支持！");
		}
		
		int minGrade = (Integer) CollectionUtils.getMinKey(gradeMaxVals);
		int maxGrade = (Integer) CollectionUtils.getMaxKey(gradeMaxVals);
		if(startGrade < minGrade){
			startGrade = minGrade;
		}
		if(endGrade > maxGrade){
			endGrade = maxGrade;
		}
		
		StringBuilder headNum = new StringBuilder("");
		for(int i= startGrade; i<= maxGrade; i++){
			headNum.append(".");
			headNum.append(gradeMaxVals.get(i));
		}
		
		String headNums = headNum.toString();
		if(headNums.startsWith(".")){
			headNums = headNums.replaceFirst(".", "");
		}
		
		return headNums;
	}
	
	/**
	 * 根据现有 编码 反推 gradeMaxVals，编码字符串的首个数字默认为等级1，正则表达式为默认 #.#
	 * @param headNum 符合 pattern 正则表达式的 编码字符串
	 * @return
	 */
	public Map<Integer, Integer> parse(String headNum){
		return parse(headNum, 1, this.pattern);
	}
	
	/**
	 * 根据现有 编码 反推 gradeMaxVals
	 * @param headNum 符合 pattern 正则表达式的 编码字符串
	 * @param startGrade 编码字符串中首个数字所代表的编码等级
	 * @param pattern 编码格式的正则表达式
	 * @return
	 */
	public Map<Integer, Integer> parse(String headNum, int startGrade, String pattern){
		if(headNum ==null || "".equals(headNum.trim())){
			return null;
		}
		
		if(startGrade <1){
			throw new ArrayIndexOutOfBoundsException("起始等级应当 ≥1！");
		}
		
		String[] numStrArray = headNum.split(".");
		if(numStrArray.length <1){
			throw new ArrayIndexOutOfBoundsException("编码字符无效！");
		}
		
		if(pattern != "#.#"){
			throw new ArrayIndexOutOfBoundsException("非标准格式的暂未支持！");
		}
		
		int level = startGrade + numStrArray.length - 1;
		
		Map<Integer, Integer> gradeMaxVals = new HashMap<Integer, Integer>();
		for(int i=0; i< level; i++){
			if(i< startGrade){
				gradeMaxVals.put(i+1, 0);
			}else{
				gradeMaxVals.put(i+1, new Integer(numStrArray[i-level+1]));
			}
		}
		return gradeMaxVals;
	}
	
}

