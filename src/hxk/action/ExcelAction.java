package hxk.action;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import hxk.util.WebExcelUtil;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * @author hxk
 * @description 一个简单的通过封装http请求返回excel文件的demo
 *2015年11月5日  下午2:51:37
 */
@Controller
@RequestMapping("/excel/")
public class ExcelAction {
    @RequestMapping("down")
    public void download(HttpServletRequest request,HttpServletResponse response){
	//设置返回格式
	WebExcelUtil.setHttpMessage(request, response);
	try {
	    OutputStream out = response.getOutputStream();
	    // 1、创建工作簿(WritableWorkbook)对象，
	    WritableWorkbook workbook = Workbook.createWorkbook(out);
	    
	    // 2、新建工作表(sheet)对象，并声明其属于第几页
	    WritableSheet sheet1 = workbook.createSheet("测试1", 0);
	    
            //设置标头
            List<String> headersList = setTitle(sheet1);
            
            //设置内容
            setData(sheet1, headersList);
            
            //关闭excel以及io流
            workbook.write();
            workbook.close();
            out.close();
            
	} catch (IOException e) {
	    e.printStackTrace();
	} catch (WriteException e) {
	    e.printStackTrace();
	}
    }

    /** @description	
     * @param sheet1
     * @param headersList
     * @throws WriteException
     * @throws RowsExceededException
     *2015年11月5日  下午3:53:22
     *返回类型:void	
     */
    private void setData(WritableSheet sheet1, List<String> headersList) throws WriteException, RowsExceededException {
	WritableCellFormat cellFormat = WebExcelUtil.getCellFormat();
	//模拟填写五行数据
	//i表示行数,j表示列数
	for (int i = 1; i < 5; i++) {
    	    for (int j = 0; j < headersList.size(); j++) {
    	       Label label = new Label(j,i,i+"----"+j,cellFormat);
    	       sheet1.addCell(label);
    	    }
	}
    }

    /** @description 设置标题内容	
     * @param sheet1 
     * @return
     * @throws RowsExceededException
     * @throws WriteException
     *2015年11月5日  下午3:50:09
     *返回类型:List<String>	
     */
    private List<String> setTitle(WritableSheet sheet1) throws RowsExceededException, WriteException {
	String headers = "姓名,性别,年龄,星座";
	List<String> headersList = Arrays.asList(headers.split(","));
	WebExcelUtil.appendColumn(sheet1, headersList);
	return headersList;
    }

  
}
