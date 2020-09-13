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
                    //��ӷ�Դtitle
                    subResult.add(subAddressList.get(0));
                    //��ȡ��Դ��ַ
                    String address=subAddressList.get(1);
                    if(address.equals("����")){
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
     * ��ȡexcel�е�����
     */
    private static List<List<String>> readExcel(File file) throws Exception {

        // ��������������ȡExcel
        InputStream is = new FileInputStream(file.getAbsolutePath());
        // jxl�ṩ��Workbook��
        Workbook wb = Workbook.getWorkbook(is);
        // ֻ��һ��sheet,ֱ�Ӵ���
        //����һ��Sheet����
        Sheet sheet = wb.getSheet(0);
        // �õ����е�����
        int rows = sheet.getRows();
        // ���е�����
        List<List<String>> allData = new ArrayList<List<String>>();
        // Խ����һ�� ����������
        for (int j = 1; j < rows; j++) {

            List<String> oneData = new ArrayList<String>();
            // �õ�ÿһ�еĵ�Ԫ�������
            Cell[] cells = sheet.getRow(j);
            for (int k = 0; k < cells.length; k++) {
                if (k==0 || k==8)//��ȡÿ�е�һ�к͵�9�е�����
                {
                    oneData.add(cells[k].getContents().trim());
                }

            }
            // �洢ÿһ������
            allData.add(oneData);
            // ��ӡ��ÿһ������
            //System.out.println(oneData);

        }
        return allData;

    }

    /**
     * ������д�뵽excel��
     */
    public static  void makeExcel(List<List<String>> result) {

        //��һ��������һ��workbook��Ӧһ��excel�ļ�
        HSSFWorkbook workbook = new HSSFWorkbook();
        //�ڶ�������workbook�д���һ��sheet��Ӧexcel�е�sheet
        HSSFSheet sheet = workbook.createSheet("BD-09");
        //����������sheet������ӱ�ͷ��0�У��ϰ汾��poi��sheet������������
        HSSFRow row = sheet.createRow(0);
        //���Ĳ���������Ԫ�����ñ�ͷ
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("title");
        cell = row.createCell(1);
        cell.setCellValue("address");

        //���岽��д������
        for(int i=0;i<result.size();i++) {

            List<String> oneData = result.get(i);
            HSSFRow row1 = sheet.createRow(i + 1);
            for(int j=0;j<oneData.size();j++) {

                //������Ԫ����ֵ
                row1.createCell(j).setCellValue(oneData.get(j));
            }
        }

        //���ļ����浽ָ����λ��
        try {
            FileOutputStream fos = new FileOutputStream("/Users/huangxinran/Desktop/result.xls");
            workbook.write(fos);
            System.out.println("д��ɹ�");
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
     public static String getAK(){

        String AKs[]={
                "���AK"};
        Random random = new Random();
        int n = random.nextInt(2);
        String myAK= AKs[n];
        return myAK;
     }

}
