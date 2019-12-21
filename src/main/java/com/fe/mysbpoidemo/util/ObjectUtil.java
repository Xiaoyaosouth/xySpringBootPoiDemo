package com.fe.mysbpoidemo.util;

import java.lang.reflect.*;

/**
 * 使用反射，动态处理对象的属性值
 */
public class ObjectUtil {

    /**
     * 通过class类型获取获取对应类型的值
     * @param typeClass class类型
     * @param value 值
     * @return Object
     */
    private static Object getClassTypeValue(Class<?> typeClass, Object value){
        if(typeClass == int.class  || value instanceof Integer){
            if(null == value){
                return 0;
            }
            return value;
        }else if(typeClass == short.class){
            if(null == value){
                return 0;
            }
            return value;
        }else if(typeClass == byte.class){
            if(null == value){
                return 0;
            }
            return value;
        }else if(typeClass == double.class){
            if(null == value){
                return 0;
            }
            return value;
        }else if(typeClass == long.class){
            if(null == value){
                return 0;
            }
            return value;
        }else if(typeClass == String.class){
            if(null == value){
                return "";
            }
            return value;
        }else if(typeClass == boolean.class){
            if(null == value){
                return true;
            }
            return value;
        }else {
            return typeClass.cast(value);
        }
    }

    /**
     * 根据属性，拿到set方法，并把值set到对象中
     * @param obj 对象
     * @param clazz 对象的class
     * @param filedName 需要设置值得属性
     * @param typeClass
     * @param value
     */
    public static void setValue(Object obj,Class<?> clazz,String filedName,Class<?> typeClass,Object value){
        String methodName = "set" + filedName.substring(0,1).toUpperCase()+filedName.substring(1);
        try{
            Method method =  clazz.getDeclaredMethod(methodName, new Class[]{typeClass});
            method.invoke(obj, new Object[]{getClassTypeValue(typeClass, value)});
        }catch(Exception ex){
            ex.printStackTrace();
        }
    }

}
