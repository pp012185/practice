import com.agile.api.*;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import static com.agile.api.ChangeConstants.*;

public class ReadExcel implements ICustomAction {

   //public ReadExcel(){};

    @Override
    public ActionResult doAction(IAgileSession session, INode iNode, IDataObject obj) {


        System.out.println("------START------");

        try {
            //先抓PAGE TWO 的專案名稱
            String ProjectName = null;
            ProjectName = obj.getValue(ATT_PAGE_TWO_TEXT01).toString();
            System.out.println("專案名稱: "+ProjectName);

            //利用FileInputStream讀取該路徑之檔案，".\\"為根目錄之意思
            String filepath = "C:\\ExcelFile\\Practice4.xlsx";
            FileInputStream inp = new FileInputStream(filepath);

            //利用wb承接FileInputStream所讀取的檔案
            XSSFWorkbook wb = new XSSFWorkbook(inp);       //讀取Excel
            XSSFSheet sheet = wb.getSheetAt(0);    //讀取wb內的頁面
            XSSFRow row = sheet.getRow(0);      //讀取頁面0的第二行
            // XSSFCell cell = row.getCell(3);     //讀取第二行的第三個元素
            // System.out.println(cell);                      //可直接透過println輸出cell
            // System.out.println(sheet.getPhysicalNumberOfRows());  // 6
            // System.out.println(row.getPhysicalNumberOfCells());   // 6

            int rowlength = sheet.getPhysicalNumberOfRows();       // number of row
            int columnlength = row.getPhysicalNumberOfCells();     // number of column
            String[][] file = new String[rowlength][columnlength]; // 用String存excel的資料

            // excel data => array
            for(int i = 0; i <rowlength; i++)
            {
                XSSFRow row1 = sheet.getRow(i);
                for(int j=0; j<columnlength; j++)
                {
                    file[i][j] = row1.getCell(j).toString();
                }
            }

            /* 將讀到的Excel 印出來
            for(int i = 0; i <rowlength; i++)
            {

                for(int j=0; j<columnlength; j++)
                {
                    System.out.print(file[i][j]+ "  " );
                }
                System.out.println();
            }
            */

            // 找excel中和"專案名稱"相同的Row
            int targetrow = 0;
            for(int i = 0; i <rowlength; i++)
            {
                if(file[i][0].equals(obj.getValue(ATT_PAGE_TWO_TEXT01).toString()))
                {
                    targetrow = i;
                    System.out.println("在第 "+(targetrow+1) +" Row !");
                }
            }

            // 找所有使用者 => 放在 users裡
            IQuery q = (IQuery)session.createObject(IQuery.OBJECT_TYPE, "select * from [Users]");   // 也可用連結Database的方法
            ArrayList users = new ArrayList();
            Iterator itr = q.execute().getReferentIterator();
            while (itr.hasNext())
            {
                users.add(itr.next());
            }

            System.out.println("抓狀態");
            IChange change = (IChange) obj;
            System.out.println(change.getStatus());

            for (int i = 0; i < users.size(); i++)
            {
                IUser user =(IUser)users.get(i);        // 用IUser包
                IUser[] approver = new IUser[]{user};

                for(int j= 1;j<columnlength;j++)
                {
                    //System.out.print(file[targetrow][j] + "  ");
                    //System.out.println(user.getValue(UserConstants.ATT_GENERAL_INFO_LAST_NAME).toString());

                    if (file[targetrow][j].equals(user.getValue(UserConstants.ATT_GENERAL_INFO_LAST_NAME).toString()))
                    {
                        System.out.println("會簽加入");
                        change.addApprovers(change.getStatus(), approver, null, false, "1");
                    }

                }

                // System.out.println(user.getValue(UserConstants.ATT_GENERAL_INFO_USER_ID) + ", " + user.getValue(UserConstants.ATT_GENERAL_INFO_FIRST_NAME) + ", " + user.getValue(UserConstants.ATT_GENERAL_INFO_LAST_NAME));
            }




        } catch (APIException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


        return new ActionResult(ActionResult.STRING,"success");
    }
}


