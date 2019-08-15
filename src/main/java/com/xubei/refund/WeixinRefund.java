package com.xubei.refund;


import com.xubei.util.JsonUtil;
import com.xubei.util.OkHttpClientUtil;
import com.xubei.util.SystemUtil;
import com.xubei.util.codec.DigestUtil;
import com.xubei.util.excel.ExcelUtil;
import com.xubei.util.lang.DateUtil;
import org.junit.Test;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.TreeMap;

public class WeixinRefund {

//    交易退款
//    请求地址:	http://api.chongqilai.com/refund.shtml
//    partner:	19113401003
//    signKey:	fd91c980f7da6dc40f5bd01e321ee1ed
    public static void main(String[] args) {
        //参数
//        Date times = DateUtil.strToDate("2019-08-06 10:26:42");
//        DateUtil.dateToStr(times, "yyyyMMddHHmmss")
        String tranNo = "P2019080622111113tiu75qupgst";
        String money = "6.05";
        String times = tranNo.substring(1, 15);

        refund(tranNo, money, times);
    }

    private static void refund(String tranNo, String money, String times) {
        String refundMoney = new BigDecimal(money).multiply(new BigDecimal("100")).intValue() + "";
        String url = "http://api.chongqilai.com/refund.shtml";
        String signKey = "fd91c980f7da6dc40f5bd01e321ee1ed";
        TreeMap<String, String> paramMap = new TreeMap<>();
        paramMap.put("partner", "19113401003");
        paramMap.put("time", times);
        paramMap.put("cpCode", "19011501");
        paramMap.put("cpTranNo", tranNo);
        paramMap.put("cpRefundTranNo", SystemUtil.getRandomId());
        paramMap.put("refundPayMoney", refundMoney);
        paramMap.put("payMoney", refundMoney);
//        paramMap.put("refundPayMoney", "605");
//        paramMap.put("payMoney", "1210");
        paramMap.put("clientIp", "59.172.180.111");

        String sign = getSign(paramMap, signKey);
        paramMap.put("sign", sign);
        System.out.println(paramMap);

//        OkHttpClientUtil.getInstance().formPost(url, null, paramMap);
    }

    @Test
    public void test() throws IOException {
        List<String[]> list = ExcelUtil.read("D://190809.xlsx");
        System.out.println(list.size());
        for (String[] strArr : list) {
            String tranNo = strArr[0];
            String money = strArr[1];
            if (SystemUtil.isEmpty(tranNo)) {
                continue;
            }
            String times = tranNo.substring(1, 15);
            refund(tranNo, money, times);
            System.out.println("==========");
        }
    }

    private static String getSign(TreeMap<String, String> paramMap, String signKey) {

        StringBuilder sb = new StringBuilder();
        paramMap.forEach((k, v) -> sb.append(v).append("|"));
        sb.append(signKey);

        String signOne = DigestUtil.md5Hex(sb.toString());
        String sign = DigestUtil.md5Hex(signOne + "|" + signKey);
        return sign;
    }
}
