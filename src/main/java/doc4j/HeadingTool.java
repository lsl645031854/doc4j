/**
 * Project:Demo
 * File:HeadingTool.java
 * Copyright 2004-2018 Homolo Co., Ltd. All rights reserved.
 */
package doc4j;

import java.util.HashMap;
import java.util.Map;

/**
 * word 文档中标题样式、编码生成等相关的工具类
 */
public class HeadingTool {
	
	/**
	 * 等级与该等级最大值Map，其中key是等级，value是该等级当前最大值
	 */
	Map<Integer, Integer> gradeMaxVals;
	
	/**
	 * 等级与该等级的统一样式Map，其中key是等级，value是该等级的样式类
	 */
	Map<Integer, HeadingStyle> gradeHeadingStyles;
	
	public Map<Integer, Integer> getGradeMaxVals(){
		return gradeMaxVals;
	}
	
	public void setGradeMaxVals(Map<Integer, Integer> gradeMaxVals) {
		this.gradeMaxVals = gradeMaxVals;
	}

	public Map<Integer, HeadingStyle> getGradeHeadingStyles() {
		return gradeHeadingStyles;
	}

	public void setGradeHeadingStyles(Map<Integer, HeadingStyle> gradeHeadingStyles) {
		this.gradeHeadingStyles = gradeHeadingStyles;
	}
	
	public HeadingTool() {
		super();
		this.gradeMaxVals = initGradeMaxVals();
		this.gradeHeadingStyles = initGradeHeadingStyles();
	}

	public HeadingTool(Map<Integer, Integer> gradeMaxVals,
			Map<Integer, HeadingStyle> gradeHeadingStyles) {
		super();
		this.gradeMaxVals = gradeMaxVals;
		this.gradeHeadingStyles = gradeHeadingStyles;
	}

	/**
	 * 初始化 gradeMaxVals
	 * @return
	 */
	public Map<Integer, Integer> initGradeMaxVals(){
		Map<Integer, Integer> _gradeMaxVals = new HashMap<Integer, Integer>();
		_gradeMaxVals.put(1, 0);
		return _gradeMaxVals;
	}
	
	/**
	 * 初始化 gradeHeadingStyles
	 * @return
	 */
	public Map<Integer, HeadingStyle> initGradeHeadingStyles(){
		Map<Integer, HeadingStyle> _gradeHeadingStyles = new HashMap<Integer, HeadingStyle>();
		HeadingStyle headingStyle1 = new HeadingStyle(1, "1");//1级标题应当声明1，其他等级方可声明Heading*
		_gradeHeadingStyles.put(1, headingStyle1);
		
		HeadingStyle headingStyle2 = new HeadingStyle(2, "Heading2");
		_gradeHeadingStyles.put(2, headingStyle2);
		
		HeadingStyle headingStyle3 = new HeadingStyle(3, "Heading3");
		_gradeHeadingStyles.put(3, headingStyle3);
		
		HeadingStyle headingStyle4 = new HeadingStyle(4, "Heading4");
		_gradeHeadingStyles.put(4, headingStyle4);
		
		HeadingStyle headingStyle5 = new HeadingStyle(5, "Heading5");
		_gradeHeadingStyles.put(5, headingStyle5);
		
		HeadingStyle headingStyle6 = new HeadingStyle(6, "Heading6");
		_gradeHeadingStyles.put(6, headingStyle6);
		
		HeadingStyle headingStyle7 = new HeadingStyle(7, "Heading7");
		_gradeHeadingStyles.put(7, headingStyle7);
		
		HeadingStyle headingStyle8 = new HeadingStyle(8, "Heading8");
		_gradeHeadingStyles.put(8, headingStyle8);
		
		HeadingStyle headingStyle9 = new HeadingStyle(9, "Heading9");
		_gradeHeadingStyles.put(9, headingStyle9);
		return _gradeHeadingStyles;
	}
	
	/**
	 * 以指定样式生成当前标题组下的特定自增等级的标题编码
	 * @param increGrade 特定自增等级
	 * @param headingFormat 指定的标题编码样式
	 * @return
	 */
	public String getAutoHeadingNum(int increGrade, HeadingFormat headingFormat){
		return getAutoHeadingNum( increGrade, headingFormat, this.gradeMaxVals);
	}
	
	/**
	 * 以指定样式生成指定标题组下的特定自增等级的标题编码
	 * @param increGrade 特定自增等级
	 * @param headingFormat 指定的标题编码样式
	 * @param _gradeMaxVals 指定标题组
	 * @return
	 */
	public static String getAutoHeadingNum(int increGrade, HeadingFormat headingFormat,Map<Integer, Integer> _gradeMaxVals){
		selfIncreGradeMaxVals(_gradeMaxVals, increGrade);
		return headingFormat.format(_gradeMaxVals, increGrade);
	}
	
	/**
	 * 从指定等级上自增，并更新 gradeMaxVals
	 * @param increGrade 指定的等级
	 */
	public void selfIncreGradeMaxVals(int increGrade){
		selfIncreGradeMaxVals(this.gradeMaxVals, increGrade);
	}
	
	/**
	 * 从指定等级上自增，并更新 gradeMaxVals
	 * @param _gradeMaxVals 原始 等级与该等级最大值Map，其中key是等级，value是该等级当前最大值
	 * @param increGrade 指定的等级
	 */
	public static void selfIncreGradeMaxVals(Map<Integer, Integer> _gradeMaxVals, int increGrade) {
		int minGrade = (Integer) CollectionUtils.getMinKey(_gradeMaxVals);
		int maxGrade = (Integer) CollectionUtils.getMaxKey(_gradeMaxVals);
		if(increGrade <minGrade){
			throw new ArrayIndexOutOfBoundsException("自增等级应当 ≥该Map最小key！");
		}
		
		if(maxGrade < increGrade){//增加一个子级标题
			_gradeMaxVals.put(increGrade, 1);
		}else{
			//新加父级标题
			for(int i = maxGrade; i>increGrade ; i--){
				_gradeMaxVals.remove(i);
			}
			
			Integer currentValue = _gradeMaxVals.get(increGrade);
			if(currentValue ==null){
				currentValue =0;
			}

			_gradeMaxVals.put(increGrade, currentValue + 1);
		}
	}
	
}

