package GeoCoder;


import java.io.*;
import java.net.URL;
import java.net.URLEncoder;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class GeoCoder {
    public static void main(String[] args){
        makeExcel(geoCoder());
    }

    public static  List<List<String>> geoCoder(){

        try{
            File file=new File("/Users/huangxinran/Desktop/test.xls");
            String key = getAK();
            List<List<String>> addressList = readExcel(file);
            List<List<String>> result= new ArrayList<List<String>>();

            //
            for(int i=0;i<addressList.size();i++) {
                int count=0;
                while(true){
                    List<String> subResult = new ArrayList< String>() ;
                    List<String> subAddressList=addressList.get(i);
                    //添加房源title
                    subResult.add(subAddressList.get(0));
                    //获取房源地址
                    String address=subAddressList.get(1);
                    if(address.equals("海外")){
                        String lng = "135";
                        String lat = "0";
                        String precise = "";
                        String confidence = "";

                        subResult.add(lng);
                        subResult.add(lat);
                        subResult.add(precise);
                        subResult.add(confidence);
                        result.add(subResult);
                        System.out.println("succuess:"+i);
                        break;
                    }
                    address = URLEncoder.encode(address,"UTF-8");
                    URL resjson = new URL("http://api.map.baidu.com/geocoding/v3/?address="
                            +address+"&output=json&ak="+key+"&callback=showLocation");
                    BufferedReader in = null;
                    if(resjson.openStream()!=null){
                        in = new BufferedReader(new InputStreamReader(resjson.openStream()));
                    }

                    String res;
                    StringBuilder sb = new StringBuilder("");
                    while ((res=in.readLine())!=null) {

                        sb.append(res.trim());
                    }

                    in.close();
                    String str = sb.toString();
                    //System.out.println("return json:"+str);


                    if(str!=null) {
                        int lngStart = str.indexOf("lng\":");
                        int lngEnd = str.indexOf(",\"lat");
                        int latEnd = str.indexOf("},\"precise");
                        int preciseEnd = str.indexOf(",\"confidence");
                        int confidenceEnd = str.indexOf(",\"level");
                        if (lngStart > 0 && lngEnd > 0 && latEnd > 0) {
                            String lng = str.substring(lngStart + 5, lngEnd);
                            String lat = str.substring(lngEnd + 7, latEnd);
                            String precise = str.substring(latEnd + 12, preciseEnd);
                            String confidence = str.substring(preciseEnd + 14, confidenceEnd);

                            subResult.add(lng);
                            subResult.add(lat);
                            subResult.add(precise);
                            subResult.add(confidence);
                            result.add(subResult);
                            System.out.println("succuess:"+i);
                            break;
                        }else if(count==10){
                            String lng = "";
                            String lat = "";
                            String precise = "";
                            String confidence = "";

                            subResult.add(lng);
                            subResult.add(lat);
                            subResult.add(precise);
                            subResult.add(confidence);
                            result.add(subResult);
                            System.out.println("succuess:"+i);
                            break;
                        }

                    }
                    count++;
                    Thread.sleep(300);
                }



            }
            return result;
        }catch(Exception e){
            e.printStackTrace();
        }
        return null;

    }

    /**
     *
     * 读取excel中的数据
     */
    private static List<List<String>> readExcel(File file) throws Exception {

        // 创建输入流，读取Excel
        InputStream is = new FileInputStream(file.getAbsolutePath());
        // jxl提供的Workbook类
        Workbook wb = Workbook.getWorkbook(is);
        // 只有一个sheet,直接处理
        //创建一个Sheet对象
        Sheet sheet = wb.getSheet(0);
        // 得到所有的行数
        int rows = sheet.getRows();
        // 所有的数据
        List<List<String>> allData = new ArrayList<List<String>>();
        // 越过第一行 它是列名称
        for (int j = 1; j < rows; j++) {

            List<String> oneData = new ArrayList<String>();
            // 得到每一行的单元格的数据
            Cell[] cells = sheet.getRow(j);
            for (int k = 0; k < cells.length; k++) {
                if (k==0 || k==8)//获取每行第一列和第9列的数据
                {
                    oneData.add(cells[k].getContents().trim());
                }

            }
            // 存储每一条数据
            allData.add(oneData);
            // 打印出每一条数据
            //System.out.println(oneData);

        }
        return allData;

    }

    /**
     * 将数据写入到excel中
     */
    public static  void makeExcel(List<List<String>> result) {

        //第一步，创建一个workbook对应一个excel文件
        HSSFWorkbook workbook = new HSSFWorkbook();
        //第二部，在workbook中创建一个sheet对应excel中的sheet
        HSSFSheet sheet = workbook.createSheet("BD-09");
        //第三部，在sheet表中添加表头第0行，老版本的poi对sheet的行列有限制
        HSSFRow row = sheet.createRow(0);
        //第四步，创建单元格，设置表头
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("title");
        cell = row.createCell(1);
        cell.setCellValue("address");

        //第五步，写入数据
        for(int i=0;i<result.size();i++) {

            List<String> oneData = result.get(i);
            HSSFRow row1 = sheet.createRow(i + 1);
            for(int j=0;j<oneData.size();j++) {

                //创建单元格设值
                row1.createCell(j).setCellValue(oneData.get(j));
            }
        }

        //将文件保存到指定的位置
        try {
            FileOutputStream fos = new FileOutputStream("/Users/huangxinran/Desktop/result.xls");
            workbook.write(fos);
            System.out.println("写入成功");
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
     public static String getAK(){

        String AKs[]={
                "你的AK"};
        Random random = new Random();
        int n = random.nextInt(2);
        String myAK= AKs[n];
        return myAK;
     }

}
