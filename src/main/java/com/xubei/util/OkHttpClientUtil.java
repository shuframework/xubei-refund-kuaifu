package com.xubei.util;

import com.xubei.enums.RequestMethod;
import okhttp3.FormBody;
import okhttp3.MediaType;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.RequestBody;
import okhttp3.Response;

import java.io.IOException;
import java.util.Map;
import java.util.concurrent.TimeUnit;

/**
 * 基于okhttp3 实现的客户端http请求，可支持异步请求
 * execute的方法是同步方法; enqueue的方法是异步方法;
 * 暂时加入的都是同步方法,只支持http
 *
 * @author shuheng
 */
public class OkHttpClientUtil {

    //时间单位都是秒
    public final static int READ_TIMEOUT = 100;
    public final static int CONNECT_TIMEOUT = 60;
    public final static int WRITE_TIMEOUT = 60;

    public static final MediaType JSON = MediaType.parse("application/json; charset=utf-8");
    private OkHttpClient okHttpClient;

    private OkHttpClientUtil() {
        okHttpClient = new OkHttpClient.Builder()
                //读取超时
                .readTimeout(READ_TIMEOUT, TimeUnit.SECONDS)
                //连接超时
                .connectTimeout(CONNECT_TIMEOUT, TimeUnit.SECONDS)
                //写入超时
                .writeTimeout(WRITE_TIMEOUT, TimeUnit.SECONDS)
                .build();
    }

    //内部类
    private static class InnerHttpClass {
        private static OkHttpClientUtil instance = new OkHttpClientUtil();
    }

    /**
     * 单例模式获取NetUtils
     *
     * @return
     */
    public static OkHttpClientUtil getInstance() {
        return InnerHttpClass.instance;
    }


    /**
     * 无需token的 get 请求
     * @param url
     * @return
     */
    public String get(String url) {
        return get(url, null);
    }

    //get 请求
    public String get(String url, String token) {
        return urlRequest(url, token, RequestMethod.GET);
    }

    //delete 请求
    public String delete(String url, String token) {
        return urlRequest(url, token, RequestMethod.DELETE);
    }

    // url的请求，即参数在url上
    private String urlRequest(String url, String token, RequestMethod method) {
        Request.Builder builder = new Request.Builder().url(url);
        if (RequestMethod.GET.equals(method)){
            builder.get();
        }
        if (RequestMethod.DELETE.equals(method)){
            builder.delete();
        }
        if (token != null) {
            builder.addHeader("Authorization", token);
        }
        Request request = builder.build();
        return getString(request);
    }


    /**
     * 无需token的post请求
     * @param url
     * @param json
     * @return
     */
    public String post(String url, String json) {
        return post(url, null, json);
    }

    public String post(String url, String token, String json) {
        return jsonRequest(url, token, json, RequestMethod.POST);
    }

    public String put(String url, String token, String json) {
        return jsonRequest(url, token, json, RequestMethod.PUT);
    }

    // RequestMethod 可以是自己弄的一个枚举，这里用的spring的
    // json请求， 即参数是json格式的
    private String jsonRequest(String url, String token, String json, RequestMethod method) {
        RequestBody body = RequestBody.create(JSON, json);
        Request.Builder builder = new Request.Builder().url(url);
        if (RequestMethod.POST.equals(method)){
            builder.post(body);
        }
        if (RequestMethod.PUT.equals(method)){
            builder.put(body);
        }
        if (RequestMethod.DELETE.equals(method)){
            builder.delete(body);
        }
        if (token != null) {
            builder.addHeader("Authorization", token);
        }
        Request request = builder.build();
        return getString(request);
    }

    public String formPost(String url, String token, Map<String, String> params) {
        FormBody.Builder builder = new FormBody.Builder();
        params.forEach ((k, v) -> {
            builder.add(k, v);
        });
        return formPost(url, token, builder.build());
    }

    //form post请求
    // new FormBody.Builder().add().add().build();
    public String formPost(String url, String token, FormBody formBody) {
        Request.Builder builder = new Request.Builder().url(url).method("POST", formBody);
        if (token != null) {
            builder.addHeader("Authorization", token);
        }
        Request request = builder.build();
        return getString(request);
    }

    private String getString(Request request) {
        try {
            Response response = okHttpClient.newCall(request).execute();
            String str = response.body().string();
            System.out.println("okhttp3请求返回结果" + str);
            return str;
        } catch (IOException e) {
            System.out.println("异常信息：" + e.getMessage());
        }
        return null;
    }

}
