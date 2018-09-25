package com.zy.easyexcelutilexample.utils;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @author Han.Sun
 * @since 2017年07月28日
 */

@ApiModel(description = "响应结果")
public class ResponseMessageUtils<T> implements Serializable {
    private static final long serialVersionUID = 1L;
    /**
     * 是否成功
     */
    private boolean success;

    protected String message;

    protected T data;

    protected int code;

    private Long timestamp;

    /**
     * 过滤字段：指定需要序列化的字段
     */
    private transient Map<Class<?>, Set<String>> includes;

    /**
     * 过滤字段：指定不需要序列化的字段
     */
    private transient Map<Class<?>, Set<String>> excludes;

    public static <T> ResponseMessageUtils<T> error(String message) {
        return error(message, 500, null);
    }

    public static <T> ResponseMessageUtils<T> error(T data) {
        return error(null, 500, data);
    }

    public static <T> ResponseMessageUtils<T> error(String message, int code) {
        return error(message, code, null);
    }

    public static <T> ResponseMessageUtils<T> error(String message, int code, T data) {
        ResponseMessageUtils<T> msg = new ResponseMessageUtils<>();
        msg.setMessage(message);
        msg.setCode(code);
        msg.setData(data);
        msg.setSuccess(false);
        return msg.putTimeStamp();
    }

    public static <T> ResponseMessageUtils<T> error(ErrorCodeEnumUtils codeEnum) {
        return error(codeEnum.getMessage(),codeEnum.getCode());
    }

    public static <T> ResponseMessageUtils<T> created(T data) {
        return new ResponseMessageUtils<T>()
                .data(data)
                .success(true)
                .putTimeStamp()
                .code(201);
    }

    public static <T> ResponseMessageUtils<T> ok() {
        return ok(null);
    }

    private ResponseMessageUtils<T> putTimeStamp() {
        this.timestamp = System.currentTimeMillis();
        return this;
    }

    public static <T> ResponseMessageUtils<T> ok(T data) {
        return new ResponseMessageUtils<T>()
                .data(data)
                .success(true)
                .putTimeStamp()
                .code(200);
    }

    public ResponseMessageUtils<T> data(T data) {
        this.data = data;
        return this;
    }

    public ResponseMessageUtils<T> success(boolean success) {
        this.success = success;
        return this;
    }

    public ResponseMessageUtils() {

    }

    public ResponseMessageUtils<T> include(Class<?> type, String... fields) {
        return include(type, Arrays.asList(fields));
    }

    public ResponseMessageUtils<T> include(Class<?> type, Collection<String> fields) {
        if (includes == null) {
            includes = new HashMap<>();
        }
        if (fields == null || fields.isEmpty()) {
            return this;
        }
        fields.forEach(field -> {
            if (field.contains(".")) {
                String tmp[] = field.split("[.]", 2);
                try {
                    Field field1 = type.getDeclaredField(tmp[0]);
                    if (field1 != null) {
                        include(field1.getType(), tmp[1]);
                    }
                } catch (Throwable e) {
                }
            } else {
                getStringListFromMap(includes, type).add(field);
            }
        });
        return this;
    }

    public ResponseMessageUtils<T> exclude(Class type, Collection<String> fields) {
        if (excludes == null) {
            excludes = new HashMap<>();
        }
        if (fields == null || fields.isEmpty()) {
            return this;
        }
        fields.forEach(field -> {
            if (field.contains(".")) {
                String tmp[] = field.split("[.]", 2);
                try {
                    Field field1 = type.getDeclaredField(tmp[0]);
                    if (field1 != null) {
                        exclude(field1.getType(), tmp[1]);
                    }
                } catch (Throwable e) {
                }
            } else {
                getStringListFromMap(excludes, type).add(field);
            }
        });
        return this;
    }

    public ResponseMessageUtils<T> exclude(Collection<String> fields) {
        if (excludes == null) {
            excludes = new HashMap<>();
        }
        if (fields == null || fields.isEmpty()) {
            return this;
        }
        Class type;
        if (getData() != null) {
            type = getData().getClass();
        } else {
            return this;
        }
        exclude(type, fields);
        return this;
    }

    public ResponseMessageUtils<T> include(Collection<String> fields) {
        if (includes == null) {
            includes = new HashMap<>();
        }
        if (fields == null || fields.isEmpty()) {
            return this;
        }
        Class type;
        if (getData() != null) {
            type = getData().getClass();
        } else {
            return this;
        }
        include(type, fields);
        return this;
    }

    public ResponseMessageUtils<T> exclude(Class type, String... fields) {
        return exclude(type, Arrays.asList(fields));
    }

    public ResponseMessageUtils<T> exclude(String... fields) {
        return exclude(Arrays.asList(fields));
    }

    public ResponseMessageUtils<T> include(String... fields) {
        return include(Arrays.asList(fields));
    }

    protected Set<String> getStringListFromMap(Map<Class<?>, Set<String>> map, Class type) {
        return map.computeIfAbsent(type, k -> new HashSet<>());
    }


    public ResponseMessageUtils<T> code(int code) {
        this.code = code;
        return this;
    }


    @ApiModelProperty(hidden = true)
    public Map<Class<?>, Set<String>> getExcludes() {
        return excludes;
    }

    @ApiModelProperty(hidden = true)
    public Map<Class<?>, Set<String>> getIncludes() {
        return includes;
    }




    /*********==   get set ==*********/
    @ApiModelProperty("调用结果消息")
    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }


    @ApiModelProperty(value = "状态码", required = true)
    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    @ApiModelProperty("成功时响应数据")
    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }

    @ApiModelProperty(value = "时间戳", required = true, dataType = "Long")
    public Long getTimestamp() {
        return timestamp;
    }

    public void setTimestamp(Long timestamp) {
        this.timestamp = timestamp;
    }

    @ApiModelProperty("是否成功")
    public boolean isSuccess() {
        return success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }


    @Override
    public String toString() {
        ObjectMapper mapper = new ObjectMapper();

        try {
            return mapper.writeValueAsString(this);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }
        return super.toString();
    }
}
