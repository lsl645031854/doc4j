/**
 * Project:Demo
 * File:CollectionUtils.java
 * Copyright 2004-2018 Homolo Co., Ltd. All rights reserved.
 */
package doc4j;

import java.util.Arrays;
import java.util.Map;
import java.util.Set;

/**
 * @author homolo
 * @date 2018-11-12 10:11
 */
public class CollectionUtils {
	 public static Object getMinKey(Map<Integer, Integer> map) {
		 if (map == null) {
			 return null;
		 }
		 Set<Integer> set = map.keySet();
		 Object[] obj = set.toArray();
		 Arrays.sort(obj);
		 return obj[0];
	}
	 public static Object getMaxKey(Map<Integer, Integer> map) {
		 if (map == null) {
			 return null;
		 }
		 Set<Integer> set = map.keySet();
		 Object[] obj = set.toArray();
		 Arrays.sort(obj);
		 return obj[obj.length - 1];
	 }
}
