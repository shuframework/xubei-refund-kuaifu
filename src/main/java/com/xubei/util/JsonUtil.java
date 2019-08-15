package com.xubei.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.TypeReference;
import com.alibaba.fastjson.serializer.SerializeConfig;
import com.alibaba.fastjson.serializer.SerializerFeature;
import com.alibaba.fastjson.serializer.SimpleDateFormatSerializer;

import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * 阿里巴巴的FastJson
 * com.alibaba.fastjson.annotation.JSONField
 * @JSONField(serialize=false) 在对象响应字段前加注解，这样生成的json也不包含该字段。
 * 
 * @author shuheng
 */
public class JsonUtil {

	private static SerializeConfig mapping = new SerializeConfig();
//	private static String dateFormat;
//	static {
//		dateFormat = "yyyy-MM-dd HH:mm:ss";
//	}

	/**
	 * bean转换json<br>
	 * pojo、list、map都可以, 顺序不一定是按照设置赋值的顺序 <br>
	 * 时间处理为日期戳<br>
	 * (注意如果有null值 不会被序列化)
	 * 
	 * @param obj
	 * @return
	 */
	public static String obj2JsonStr(Object obj) {
		return JSON.toJSONString(obj);
	}
	
	/**
	 * bean转换json含日期格式化，null不显示 <br>
	 * 默认的处理时间,yyyy-MM-dd HH:mm:ss <br>
	 * pojo、list、map都可以, 顺序不一定是按照设置赋值的顺序 <br>
	 * 
	 * @param obj
	 * @return
	 */
	public static String obj2JsonStrHasTime(Object obj) {
		//默认的处理时间,yyyy-MM-dd HH:mm:ss
		return JSON.toJSONString(obj, SerializerFeature.WriteDateUseDateFormat);
	}
	
	/**
	 * bean转换json<br>
	 * pojo、list、map都可以
	 * 
	 * @param obj
	 * @param dateFormat 自定义时间格式 
	 * @return
	 */
	public static String obj2JsonStr(Object obj, String dateFormat) {
		//dateFormat自定义时间格式 
		mapping.put(Date.class, new SimpleDateFormatSerializer(dateFormat));
		return JSON.toJSONString(obj, mapping);
	}
	
	/**
	 * 将json转为map
	 * 
	 * @param jsonStr
	 * @return
	 */
	@SuppressWarnings("rawtypes")
	public static Map jsonStr2Map(String jsonStr) {
		Map map = JSON.parseObject(jsonStr, Map.class);
		return map;
	}
	
	/**
	 * 将json转为T对象
	 * 
	 * @param jsonStr
	 * @return 
	 * @return
	 */
	public static <T> T jsonStr2Obj(String jsonStr,Class<T> clazz) {
		T t = JSON.parseObject(jsonStr, clazz);
		return t;
	}
	
	/**
	 * 将json转为list<Map>
	 * 
	 * @param jsonStr
	 * @return
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static List<Map> jsonStr2List(String jsonStr) {
		List<Map> list = JSON.parseObject(jsonStr, new TypeReference<List>(){});
		return list;
	}
	
	/**
	 * 将json转为list<T> 对象
	 * 
	 * @param jsonStr
	 * @param clazz
	 * @return
	 */
	public static <T> List<T> jsonStr2List(String jsonStr, Class<T> clazz) {
		List<T> list = JSON.parseArray(jsonStr, clazz);
		return list;
	}
	
}
