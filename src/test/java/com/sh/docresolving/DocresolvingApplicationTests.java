package com.sh.docresolving;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.NameValuePair;
import org.apache.http.StatusLine;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.utils.HttpClientUtils;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.util.EntityUtils;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import javax.annotation.PostConstruct;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.concurrent.CountDownLatch;

@RunWith(SpringRunner.class)
@SpringBootTest
public class DocresolvingApplicationTests {

    @Test
    public void contextLoads() {
    }

   /* @Test
    public void HSSFWORKBOOKTEST() throws Exception{
        String fileIn = "sample1/download.xlsx";
        String uri = this.getClass().getResource(fileIn).getPath();
        String fileOut = uri.replaceAll(".xls$|.xlsx$",".pdf");
        ExcelToPdf.convert(uri,fileOut);
    }

    @Test
    public void HSSFWORKBOOKTESTss() throws Exception{
        String fileIn = "F:\\docresolving\\target\\test-classes\\com\\sh\\docresolving\\sample1\\download.xlsx";
        *//*String uri = this.getClass().getResource(fileIn).getPath();
        System.out.println(fileIn);*//*
        String fileOut = fileIn.replaceAll(".xls$|.xlsx$",".pdf");
        PrintSetup printSetup = new PrintSetup();
        printSetup.put("sheet1",true);
        printSetup.put("sheet2",false);
        printSetup.put("sheet3",true);
        printSetup.put("Sheet5",false);
        Excel2Pdf.excel2Pdf(fileIn,fileOut,printSetup);
    }

    @Test
    public void sss() throws Exception{
        String fontName = "arial";
        String ttfDir =  Thread.currentThread().getContextClassLoader().getResource("ttf").getPath();
        ttfDir = replaceFirstLineChar(ttfDir);
        String fileDir = ttfDir+File.separator+fontName+".ttf";
        File file = new File(fileDir);
        Font font = java.awt.Font.createFont(0,file);
        String aa = font.getFontName();
        String bb = font.getName();
        System.out.println(aa+":"+bb);
        boolean a = file.exists();
        System.out.println(a);
    }

    @Test
    public void printOutFontFileName() throws Exception{
        String folderDir = Thread.currentThread().getContextClassLoader().getResource("ttf").getPath();
        folderDir = replaceFirstLineChar(folderDir);
        File folder = new File(folderDir);
        String[] filelist = folder.list();
        System.out.println(filelist.length);
        for (int i = 0; i < filelist.length; i++) {
            File ttfFile = new File(folderDir + "\\" + filelist[i]);
            if(!ttfFile.getName().contains("ttf")) continue;
            Font font = java.awt.Font.createFont(0,ttfFile);
            System.out.println(ttfFile.getName()+":"+font.getName());
        }
    }

    public String replaceFirstLineChar(String str){
        //BaseFont.create

        return str.substring(1);
    }

    @Test
    public void ssss(){
        getGroupNameByUrl("http://www.shouhouzn.net/group1/M00/00/17/rBGmcV1NOD2AcmuiAAG3t5XrVbc747.pdf");
    }

    private String getGroupNameByUrl(String url){
        String infoPath = url.substring(url.indexOf("group"));
        String groupName = infoPath.substring(0,infoPath.indexOf("/"));
        String path = infoPath.substring(infoPath.indexOf("/")+1);
        return url;
    }*/

    public static final int THREAD_NUM = 300;

    String[] strs = new String[]{
            "http://localhost:9227/docresolving/excel-convert/excel-to-pdf?file=http%3A%2F%2Fwww.shouhouzn.net%2Fgroup1%2FM00%2F00%2F1A%2FrBGmcV1cnuSAOyRGAACmXtRYglA90.xlsx&json=%7B%22sheet6%22%3Afalse%2C%22pageNumStart%22%3A2%2C%22headerStart%22%3A2%2C%22leftHeader%22%3A%22%E9%87%8D%E5%BA%86%E7%B1%B3%E8%88%9F%E6%A3%80%E6%B5%8B%E8%81%94%E5%8F%91%E7%A7%91%E6%8A%80%E5%85%AC%E5%8F%B8%22%2C%22rightHeader%22%3A%22M190057%22%7D",
            "http://localhost:9227/docresolving/excel-convert/excel-to-pdf?file=http%3A%2F%2Fwww.shouhouzn.net%2Fgroup1%2FM00%2F00%2F1C%2FrBGmcV1ctTuACefNAACmXtRYglA57.xlsx&json=%7B%22sheet6%22%3Afalse%2C%22pageNumStart%22%3A2%2C%22headerStart%22%3A2%2C%22leftHeader%22%3A%22%E9%87%8D%E5%BA%86%E7%B1%B3%E8%88%9F%E6%A3%80%E6%B5%8B%E8%81%94%E5%8F%91%E7%A7%91%E6%8A%80%E5%85%AC%E5%8F%B8%22%2C%22rightHeader%22%3A%22M190057%22%7D",
            "http://localhost:9227/docresolving/excel-convert/excel-to-pdf?file=http%3A%2F%2Fwww.shouhouzn.net%2Fgroup1%2FM00%2F00%2F1C%2FrBGmcV1ctVKAEPIPAACmXtRYglA22.xlsx&json=%7B%22sheet6%22%3Afalse%2C%22pageNumStart%22%3A2%2C%22headerStart%22%3A2%2C%22leftHeader%22%3A%22%E9%87%8D%E5%BA%86%E7%B1%B3%E8%88%9F%E6%A3%80%E6%B5%8B%E8%81%94%E5%8F%91%E7%A7%91%E6%8A%80%E5%85%AC%E5%8F%B8%22%2C%22rightHeader%22%3A%22M190057%22%7D",
            "http://localhost:9227/docresolving/excel-convert/excel-to-pdf?file=http%3A%2F%2Fwww.shouhouzn.net%2Fgroup1%2FM00%2F00%2F1C%2FrBGmcV1ctWqAA5AoAACmXtRYglA33.xlsx&json=%7B%22sheet6%22%3Afalse%2C%22pageNumStart%22%3A2%2C%22headerStart%22%3A2%2C%22leftHeader%22%3A%22%E9%87%8D%E5%BA%86%E7%B1%B3%E8%88%9F%E6%A3%80%E6%B5%8B%E8%81%94%E5%8F%91%E7%A7%91%E6%8A%80%E5%85%AC%E5%8F%B8%22%2C%22rightHeader%22%3A%22M190057%22%7D",
            "http://localhost:9227/docresolving/excel-convert/excel-to-pdf?file=http%3A%2F%2Fwww.shouhouzn.net%2Fgroup1%2FM00%2F00%2F1C%2FrBGmcV1ctYyACfaoAACmXtRYglA88.xlsx&json=%7B%22sheet6%22%3Afalse%2C%22pageNumStart%22%3A2%2C%22headerStart%22%3A2%2C%22leftHeader%22%3A%22%E9%87%8D%E5%BA%86%E7%B1%B3%E8%88%9F%E6%A3%80%E6%B5%8B%E8%81%94%E5%8F%91%E7%A7%91%E6%8A%80%E5%85%AC%E5%8F%B8%22%2C%22rightHeader%22%3A%22M190057%22%7D"
    };

    /**
     * 开始时间
     */
    private static long startTime = 0L;

    @PostConstruct
    public void init() {
        try {

            startTime = System.currentTimeMillis();
            System.out.println("CountDownLatch started at: " + startTime);

            // 初始化计数器为1
            CountDownLatch countDownLatch = new CountDownLatch(1);

            for (int i = 0; i < THREAD_NUM; i ++) {
                new Thread(new Run(countDownLatch)).start();
            }

            // 启动多个线程
            countDownLatch.countDown();

        } catch (Exception e) {
            System.out.println("Exception: " + e);
        }
    }

    /**
     * 线程类
     */
    private class Run implements Runnable {
        private final CountDownLatch startLatch;

        public Run(CountDownLatch startLatch) {
            this.startLatch = startLatch;
        }

        @Override
        public void run() {
            try {
                // 线程等待
                startLatch.await();

                // 执行操作
                /**
                 这里调用你要测试的接口
                 */
                CloseableHttpClient httpClient = HttpClients.createDefault();
                int random =  new Random().nextInt(5);
                HttpPost httpPost = new HttpPost(strs[random]);// 创建httpPost
                httpPost.setHeader("Content-Type", "application/json");
                List<NameValuePair> formParams = new ArrayList<>();
                formParams.add(new BasicNameValuePair("json","{\"sheet6\":false,\"pageNumStart\":2,\"headerStart\":2,\"leftHeader\":\"重庆米舟检测联发科技公司\",\"rightHeader\":\"M190057\"}"));
                formParams.add(new BasicNameValuePair("file","http://www.shouhouzn.net/group1/M00/00/1A/rBGmcV1cnuSAOyRGAACmXtRYglA90.xlsx"));
                HttpEntity entity1 =
                        new UrlEncodedFormEntity(formParams, "utf-8");
                httpPost.setEntity(entity1);
                HttpResponse response = httpClient.execute(httpPost);
                long endTime = System.currentTimeMillis();
                StatusLine statusLine = response.getStatusLine();
                int code = statusLine.getStatusCode();
                if (code == 200) {
                    HttpEntity entity2 = response.getEntity();
                    System.out.println("返回:"+EntityUtils.toString(entity2, "utf-8"));
                }
                System.out.println(Thread.currentThread().getName() + " ended at: " + endTime + ", cost: " + (endTime - startTime) + " ms.");
            } catch (Exception e) {
                e.printStackTrace();
            }

        }
    }

    @Test
    public static void main(String[] args) {
        DocresolvingApplicationTests apiTest = new DocresolvingApplicationTests();
        apiTest.init();
    }
}
