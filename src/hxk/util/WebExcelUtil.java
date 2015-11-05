package hxk.util;

import java.io.UnsupportedEncodingException;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * @author hxk
 * @description
 *2015年11月5日  下午3:56:06
 */
public class WebExcelUtil {
    /** @description 列值的单元格格式	
     * @return
     * @throws WriteException
     *2015年11月5日  下午3:32:44
     *返回类型:WritableCellFormat	
     */
    public static WritableCellFormat getCellFormat() throws WriteException {
	WritableFont cellFont = new WritableFont(WritableFont.ARIAL, 9);
	WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
	cellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
	cellFormat.setAlignment(Alignment.CENTRE);
	cellFormat.setWrap(true);
	return cellFormat;
    }
    
     

    /** @description 设置返回的http协议为excel	
     * @param request
     * @param response
     *2015年11月5日  下午2:56:27
     *返回类型:void	
     */
    public static void setHttpMessage(HttpServletRequest request,HttpServletResponse response) {
	try {
	    response.reset();// 重置
	    request.setCharacterEncoding("UTF-8");
	    String agent = request.getHeader("USER-AGENT"); 
	    String title = "测试excel下载";
	    String downLoadName = null;
	    //IE11浏览器的user-agent使用MSIE容易识别为firefox  导致出错
	    //区分浏览器   修改之后如下
	    if(null != agent && -1!= agent.toLowerCase().indexOf("firefox")){//是firefox
		 downLoadName = new String(title.getBytes("UTF-8"),"iso-8859-1");     
	    }else if(null != agent && -1 != agent.toUpperCase().indexOf("CHROME")){//chrome
		 downLoadName = java.net.URLEncoder.encode(title, "UTF-8");   
	    }else{//IE
		 downLoadName = java.net.URLEncoder.encode(title, "UTF-8");   
	    }
	    response.setHeader("Content-Disposition", "attachment;filename=" + downLoadName + ".xls");// 表示以附件形式可下载
	    response.setContentType("application/vnd.ms-excel; charset=utf-8");// 设置下载格式为EXCEL
	} catch (UnsupportedEncodingException e) {
	    e.printStackTrace();
	}
    }
    
    
    /**
     * @description 设置标题格式	
     * @return
     * @throws WriteException
     *2015年11月5日  下午3:42:37
     *返回类型:WritableCellFormat
     */
    public static WritableCellFormat getTitleCellFormat() throws WriteException{
	WritableFont columnFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE, Colour.BLACK);
	// 列标题格式
        WritableCellFormat titleCellFormat = new WritableCellFormat(columnFont);
        titleCellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
        titleCellFormat.setAlignment(Alignment.CENTRE);
        titleCellFormat.setVerticalAlignment(VerticalAlignment.CENTRE);
        titleCellFormat.setWrap(true);
        return titleCellFormat;
    }
    
    /**
	 * Excel添加列标题
	 * 
	 * @param sheet
	 * @param headers 标题列表
	 * @throws RowsExceededException
	 * @throws WriteException
    */
    public static void appendColumn(WritableSheet sheet, List<String> headers) throws RowsExceededException, WriteException {
    	sheet.setRowView(0, 500);//设置第一行的高度
    	if (headers.size() > 0) {
    	    for (int i = 0; i <headers.size(); i++) {
    		Label label = new Label(i, 0, headers.get(i), getTitleCellFormat());
    		//以下是设置列宽的代码	
    		//sheet.setColumnView(i, 80);
    		sheet.addCell(label);
    	    }
    	}
    }
}
