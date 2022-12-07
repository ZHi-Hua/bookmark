package com.gentlesoft.workform.word.service;

import com.gentlesoft.commons.util.Base64;
import com.gentlesoft.commons.util.UtilValidate;
import com.gentlesoft.commons.util.json.JsonUtil;
import com.gentlesoft.workform.word.util.data.BookMark;
import com.gentlesoft.workform.word.util.data.MSWordTool;
import com.gentlesoft.workform.word.util.data.Picture;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.RoundingMode;
import java.net.URL;
import java.text.NumberFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DealWithTaoHong {
    private static final Logger log = LoggerFactory.getLogger(DealWithTaoHong.class);

    public static void main(String[] args) {
        Map<String, Object> data = JsonUtil.jsonToMap("{\n" +
                "    \"版本更新#vtype\": \"default\",\n" +
                "    \"图片\": \"\",\n" +
                "    \"编辑域#fieldName\": \"EDITNESKDDR\",\n" +
                "    \"文本区#fieldName\": \"TEXTCNZOSNZD\",\n" +
                "    \"超链接\": \"\",\n" +
                "    \"日期#fieldName\": \"DATEYWYEVLL\",\n" +
                "    \"单选框#vtype\": \"default\",\n" +
                "    \"日期#vtype\": \"default\",\n" +
                "    \"大文本域#fieldName\": \"BIGTEVPGKNH\",\n" +
                "    \"附件域#fieldName\": \"FILEVURRZVF\",\n" +
                "    \"PINGLUNZUVWMV\": \"\",\n" +
                "    \"单选框#json\": \"W3sidmFsdWUiOiIxIiwiZGVzY3JpcHRpb24iOiJxIiwiZGlzcGxheVVybCI6bnVsbCwiY2hlY2tlZCI6dHJ1ZX0seyJ2YWx1ZSI6IjIiLCJkZXNjcmlwdGlvbiI6InciLCJkaXNwbGF5VXJsIjpudWxsLCJjaGVja2VkIjpmYWxzZX1d\",\n" +
                "    \"文本框#vtype\": \"default\",\n" +
                "    \"下拉框#type\": \"drop-down\",\n" +
                "    \"特殊编辑域#viewConfig\": \"eyJzaG93TW9kZml5RGF0ZSI6InRydWUiLCJzaG93SGlzdG9yeVVzZXJJbmZvIjoiZmFsc2UiLCJkYXRlUGF0dGVybiI6Inl5eXktTU0tZGQgSEg6bW06c3M6U1NTIiwic2hvd0hpc3RvcnlEYXRlIjoiZmFsc2UifQ==\",\n" +
                "    \"隐藏域#fieldName\": \"HIDDENPPVKCJ\",\n" +
                "    \"创建时间#type\": \"hidden\",\n" +
                "    \"选择域#vtype\": \"distInfo_false\",\n" +
                "    \"文本区#type\": \"textarea\",\n" +
                "    \"图片#vtype\": \"undefined\",\n" +
                "    \"版本更新\": \"MTY2MzU1MzI5MzEwMA==\",\n" +
                "    \"选择域#fieldName\": \"SEDITZYZLVNX\",\n" +
                "    \"复选框#json\": \"W3sidmFsdWUiOiIxIiwiZGVzY3JpcHRpb24iOiJhIiwiZGlzcGxheVVybCI6bnVsbCwiY2hlY2tlZCI6ZmFsc2V9LHsidmFsdWUiOiIyIiwiZGVzY3JpcHRpb24iOiJiIiwiZGlzcGxheVVybCI6bnVsbCwiY2hlY2tlZCI6dHJ1ZX1d\",\n" +
                "    \"复选框#printType\": \"noframe\",\n" +
                "    \"单选框#fieldName\": \"RADIODHYP8CEE7LW\",\n" +
                "    \"显示域\": \"\",\n" +
                "    \"按钮域#type\": \"button-field\",\n" +
                "    \"下拉框#fieldName\": \"SELECTXNMZYO\",\n" +
                "    \"版本更新#fieldName\": \"VERSION\",\n" +
                "    \"特殊编辑域#current\": \"eyJmaWVsZE5hbWUiOiJTUEVESVRQTVdHWEYiLCJzZW5kZWQiOiIwIiwidXNlck5hbWUiOiLnrqHnkIblkZgiLCJzdWJUaW1lIjoxNjYzNTUzMjkzMDAwLCJ1c2VySWQiOiJhZG1pbiIsImFzc2lnbm1lbnRJZCI6bnVsbCwiYWN0TmFtZSI6IuaZrumAmua0u+WKqCIsInN1YkRhdGVUaW1lU3RyIjpudWxsLCJzdWJEYXRlVGltZSI6IjIwMjItMDktMTkiLCJhY3REZWZJZCI6IlBhY2thZ2VfSktYWEQwVFhfV29yMV9BY3QxIiwiZGF0YUlkIjoiUTFER28zTGRNcFA5dEVXVWZhRyIsInNpZ25UeXBlIjoiMiIsImlkIjoiZG9YcFRCV3licUdQM01GZFdIMSIsInN0ckNvbnRlbnQiOiLnibnmrornvJbovpHln5/oi43kupHOvlxyXG5cdTAwM2NJTUcgIG9uZXJyb3JcdTAwM2RcdTAwMjd0aGlzLm91dGVySFRNTFx1MDAzZFwi566h55CG5ZGYXCJcdTAwMjcgc3JjXHUwMDNkXHUwMDI3aHR0cDovL2xvY2FsaG9zdDo4MDgwL3BsYXRmb3JtLy9yZXNvdXJjZXMvd29ya2Zsb3cvd29ya2Zvcm0vc2lnbl9waWMvYWRtaW4uanBnXHUwMDI3XHUwMDNlIDIwMjItMDktMTkgMDk6NTE6MDE6NDk3In0=\",\n" +
                "    \"PINGLUNZUVWMV#type\": \"suggest\",\n" +
                "    \"单选框\": \"cQ==\",\n" +
                "    \"复选框#vtype\": \"checkbox-horizontal\",\n" +
                "    \"按钮域\": \"\",\n" +
                "    \"特殊编辑域#showContent\": \"true\",\n" +
                "    \"编辑域#type\": \"edit-field\",\n" +
                "    \"附件域#type\": \"adjunct-field\",\n" +
                "    \"主键标识#type\": \"hidden\",\n" +
                "    \"单选框#printType\": \"noframe\",\n" +
                "    \"特殊编辑域\": \"\",\n" +
                "    \"主键标识#vtype\": \"default\",\n" +
                "    \"隐藏域\": \"\",\n" +
                "    \"显示域#type\": \"display\",\n" +
                "    \"日期#type\": \"date-time\",\n" +
                "    \"复选框\": \"Yg==\",\n" +
                "    \"复选框#fieldName\": \"CHECKOMXCBP\",\n" +
                "    \"特殊编辑域#type\": \"special-edit\",\n" +
                "    \"文本框#fieldName\": \"TEBOXGCRQET\",\n" +
                "    \"附件域\": \"RklMRV9NSk9QQ1JK\",\n" +
                "    \"PINGLUNZUVWMV#vtype\": \"pl\",\n" +
                "    \"超链接#vtype\": \"default\",\n" +
                "    \"按钮域#vtype\": \"default\",\n" +
                "    \"日期\": \"MjAyMi0wOS0xOQ==\",\n" +
                "    \"版本更新#type\": \"hidden\",\n" +
                "    \"主键标识#fieldName\": \"DATA_ID\",\n" +
                "    \"下拉框\": \"YWE=\",\n" +
                "    \"文本区\": \"MTIzMg==\",\n" +
                "    \"编辑域#vtype\": \"commonWord\",\n" +
                "    \"特殊编辑域#vtype\": \"sign_field_hascontent_image_onlyadd\",\n" +
                "    \"大文本域#type\": \"bigtext\",\n" +
                "    \"显示域#fieldName\": \"DISPLAYWDQMMI\",\n" +
                "    \"大文本域\": \"\",\n" +
                "    \"图片#fieldName\": \"IMGVSBPIKW\",\n" +
                "    \"创建时间#fieldName\": \"CREATE_DATE\",\n" +
                "    \"隐藏域#vtype\": \"default\",\n" +
                "    \"下拉框#vtype\": \"default\",\n" +
                "    \"大文本域#vtype\": \"simpleText\",\n" +
                "    \"特殊编辑域#fieldName\": \"SPEDITPMWGXF\",\n" +
                "    \"创建时间#vtype\": \"default\",\n" +
                "    \"编辑域\": \"\",\n" +
                "    \"文本框\": \"MTEx\",\n" +
                "    \"隐藏域#type\": \"hidden\",\n" +
                "    \"超链接#type\": \"hyperlink\",\n" +
                "    \"图片#type\": \"image-field\",\n" +
                "    \"文本框#type\": \"text\",\n" +
                "    \"文本区#vtype\": \"default\",\n" +
                "    \"超链接#fieldName\": \"AAPOIUXULFT\",\n" +
                "    \"选择域#type\": \"select-field\",\n" +
                "    \"显示域#vtype\": \"\",\n" +
                "    \"选择域\": \"\",\n" +
                "    \"创建时间\": \"MTY2MzI5MzU0OTIxMQ==\",\n" +
                "    \"特殊编辑域#json\": \"W10=\",\n" +
                "    \"按钮域#fieldName\": \"BUTJNSSMLF\",\n" +
                "    \"下拉框#json\": \"W3sidmFsdWUiOiIxIiwiZGVzY3JpcHRpb24iOiJhYSIsImRpc3BsYXlVcmwiOm51bGwsImNoZWNrZWQiOmZhbHNlfSx7InZhbHVlIjoiMiIsImRlc2NyaXB0aW9uIjoiYmIiLCJkaXNwbGF5VXJsIjpudWxsLCJjaGVja2VkIjpmYWxzZX1d\",\n" +
                "    \"主键标识\": \"UTFER28zTGRNcFA5dEVXVWZhRw==\",\n" +
                "    \"复选框#type\": \"check\",\n" +
                "    \"附件域#vtype\": \"default\",\n" +
                "    \"PINGLUNZUVWMV#fieldName\": \"PINGLUNZUVWMV\",\n" +
                "    \"单选框#type\": \"radio\"\n" +
                "}");
        MSWordTool doc = new MSWordTool();
        doc.setTemplate("E:\\文件\\新建 DOCX 文档.docx");
        DealWithTaoHong d = new DealWithTaoHong();
        d.replaceBookMarkNewBase64(data, doc);
    }
    //替换word文档模板中的书签(与后台base64相互解码)
    private void replaceBookMarkNewBase64(Map<String, Object> data, MSWordTool doc) {
        //遍历
        for (String key : data.keySet()) {
            //获得书签名称
            if (key.indexOf("#") <= 0) {
                copyToBookMark(key, data, doc);
            }
        }
    }


    private void copyToBookMark(String bookMarkName, Map<String, Object> data, MSWordTool doc) {
        String rvalue = "";
        String dataValue = (String) data.get(bookMarkName);
        //对输入域中的值进行base64解码
        if (dataValue != "") {
            rvalue = Base64.base64Decode(dataValue);
            rvalue = rvalue.replace("#", "\r\n");
        }
        String type = (String) data.get(bookMarkName + "#type");
        String vtype = (String) data.get(bookMarkName + "#vtype");
        String json = "";
        String jsonField = (String) data.get(bookMarkName + "#json");
        String value = "";
        String valueField = (String) data.get(bookMarkName + "#value");
        if (UtilValidate.isNotEmpty(jsonField)) {
            json = Base64.base64Decode(jsonField);
        }
        if (UtilValidate.isNotEmpty(valueField)) {
            value = valueField;
        }
        try {
            BookMark bkmkObj = doc.getBookMarks().getBookmark(bookMarkName);
            if (bkmkObj == null) {
                return;
            }
            boolean useDefalut = true;
            boolean useStrValue = false;
            if (useDefalut) {
                useStrValue = copyToBookMarkByDefalut(doc.getDocument(), bkmkObj, bookMarkName, type, vtype, value, rvalue, json, data);
                /*XWPFDocument doc1 = new XWPFDocument();*/
                /*XWPFTable table = doc1.insertNewTbl();
                XWPFTableRow row = table.createRow();
                row.createCell().setText("");
                row.createCell().setText("11");*/
                if (useStrValue) {

                }
            }

        } catch (Exception e) {
            log.error(e.getMessage());
            throw new RuntimeException(e.getMessage());
        }
    }

    private boolean copyToBookMarkByDefalut(XWPFDocument doc, BookMark bkmkObj, String bookMarkName, String type, String vType, String value, String rValue, String json, Map<String, Object> data) {
        try {
            String symbolValue_true = "☑"; //多选框选中样式☑■
            String symbolValue_false = "□"; //多选框未选中样式□
            String radioSymbolValue_true = "⊙"; //单选框选中样式☑■
            String radioSymbolValue_false = "○"; //单选框未选中样式□
            switch (type) {
                case "special-edit":
                    dealWithSpecialEdit(doc, bkmkObj, bookMarkName, type, vType, value, rValue, json, data);
                    return false;
                case "image-field":
                    //bkmkObj.Range.InlineShapes.AddPicture(getcontextUrl()+'/adjunct/showIcon.png?dataId='+dataId+'&adjunctId='+value);
                    //return false;
                    break;
                case "bigtext":
                  /*  var saverange = bkmkObj.Range
                    saverange.Text = rvalue.replace( / < br >/ig, "");
                    webWord.office.ActiveDocument.Bookmarks.Add(BookMarkName, saverange);*/
                    break;
                case "check":
                case "radio":
                   /* var dataValue = eval(json);
                    var printType = datas[fieldName + '#printType'] ? datas[fieldName + '#printType'] : "";*/
                    String printType = "";
                    switch (printType) {
                        case "noframe": //只打印选中的值,与之前3.0版本一样
                            return true;
                        case "frame": //新版打印样式
                           /* var HorS = vtype == "default" ? true : false;//true默认竖排,false扩展类型横排
                            if (HorS) {//竖排(创建虚拟表格,实现竖排展示)
                                var objTable_check = webWord.office.ActiveDocument.Tables.Add(bkmkObj.Range, dataValue.length, 1);//创建一个表格做换行使用
                                for (var i = 0; i < dataValue.length; i++) {
                                    var showValue = "";
                                    if (type == "check")
                                        showValue = dataValue[i].checked == true ? symbolValue_true : symbolValue_false;
                                    else if (type == "radio")
                                        showValue = dataValue[i].checked == true ? radioSymbolValue_true : radioSymbolValue_false;
                                    showValue += dataValue[i].description;
                                    objTable_check.Cell(i + 1, 1).Range.InsertAfter(showValue);
                                }
                            } else {//横排
                                var showValue = "";
                                for (var i = 0; i < dataValue.length; i++) {
                                    if (type == "check")
                                        showValue += dataValue[i].checked == true ? symbolValue_true : symbolValue_false;
                                    else if (type == "radio")
                                        showValue += dataValue[i].checked == true ? radioSymbolValue_true : radioSymbolValue_false;
                                    showValue += dataValue[i].description + "  ";
                                }
                                var saverange = bkmkObj.Range
                                saverange.Text = showValue;
                                webWord.office.ActiveDocument.Bookmarks.Add(BookMarkName, saverange);
                            }*/
                            return false;
                        default:  //可继续扩展打印样式
                            return true;
                    }
                default:
                    break;

            }
        } catch (Exception e) {

        }
        return true;
    }


    public void dealWithSpecialEdit(XWPFDocument doc, BookMark bkmkObj, String bookMarkName, String type, String vType, String value, String rValue, String json, Map<String, Object> data) throws Exception {
        List<Map<String, Object>> list = JsonUtil.jsonArrayToList(json);
        //初始化视图配置
        Map<String, Object> viewConfig = new HashMap<>();
        if (UtilValidate.isNotEmpty((String) data.get(bookMarkName + "#viewConfig"))) {
            viewConfig = JsonUtil.jsonToMap(Base64.base64Decode((String) data.get(bookMarkName + "#viewConfig")));
        }
        if (UtilValidate.isEmpty((String) viewConfig.get("showHistoryDate"))) {
            viewConfig.put("showHistoryDate", true);
        }
        if (UtilValidate.isEmpty((String) viewConfig.get("showHistoryUserInfo"))) {
            viewConfig.put("showHistoryUserInfo", true);
        }
        if (UtilValidate.isEmpty((String) viewConfig.get("showModfiyDate"))) {
            viewConfig.put("showModfiyDate", false);
        }
        if (UtilValidate.isNotEmpty((String) data.get(bookMarkName + "#current"))) {
            String currentValue = (String) data.get(bookMarkName + "#current");
            list.add(JsonUtil.jsonToMap(Base64.base64Decode(currentValue)));
        }
        boolean disCon = data.get(bookMarkName + "#showContent") == "true";
        int row = disCon ? list.size() * 2 : list.size();
        int col = 1;
        XmlCursor cursor = bkmkObj.getPara().getCTP().newCursor();
        XWPFTable table = doc.insertNewTbl(cursor);
        for (int i = 0; i < col; i++) {
            if (i > 0) {
                table.getRow(0).createCell();
            }
        }
       /* List<Object> specialList = new ArrayList<>();*/
        for (int i = 0; i < list.size(); i++) {
            for (int j = 0; j < col; j++) {
                if (disCon) {
                    //如果显示修改时间
                    if ((boolean) viewConfig.get("showModfiyDate") == true) {
                        String content = (String) list.get(i).get("strContent");
                        if (list.get(i).get("signType") == "2") {
                            //图片
                            String sreg = "/<\\s*img\\s+onerror\\s*=\\s*\\'this.outerHTML\\s*=\\\"(\\S+)\\\"\\'\\s+src\\s*=\\s*[\\'|\\\"](\\S+)[\\'|\\\"]\\s*\\/?>\\s*(.*)/i";
                            Pattern pattern = Pattern.compile(sreg);
                            Matcher matcher = pattern.matcher(content);
                            ArrayList al = new ArrayList();
                            while (matcher.find()) {
                                al.add(matcher.group(0));
                            }
                            if (table.getRow(i) == null) {
                                table.createRow();
                            }
                            table.getRow(i).getCell(j).setText(content.substring(0, content.indexOf('ξ')));
                            if (al.get(1) != null) {
                                table.createRow().getCell(j).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                                try {
                                    this.setCellImg(table.getRow(i).getCell(j), "getcontextUrl() + group[2].substring(2, group[2].length)");
                                    XWPFParagraph paragraph = table.getRow(i).getCell(j).addParagraph();
                                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                                    XWPFRun run = paragraph.getRuns().isEmpty() ? paragraph.createRun() : paragraph.getRuns().get(0);
                                    run.setText((String) al.get(2));
                                } catch (RuntimeException e) {
                                    table.createRow().getCell(j).setText("" + al.get(1) + al.get(3));
                                }
                            }
                        } else {
                            String[] contents = content.split("ξ");
                            if (table.getRow(i) == null) {
                                table.createRow();
                            }
                            if (contents.length == 2) {
                                table.getRow(i).getCell(j).setText(contents[0]);
                                table.createRow().getCell(j).setText(contents[1].trim());
                                        /*objTable.Cell(i + k, j + 1).Range.InsertAfter(contents[0]);
                                        objTable.Cell(i + k + 1, j + 1).Range.ParagraphFormat.Alignment = 2
                                        objTable.Cell(i + k + 1, j + 1).Range.InsertAfter(contents[1].trim());*/
                            } else {
                                table.getRow(i).getCell(j).setText(content);
                                //objTable.Cell(i + k, j + 1).Range.InsertAfter(content);
                            }
                        }
                    } else {
                        if (list.get(i).get("signType") != "3") {
                            if (table.getRow(i) == null) {
                                table.createRow();
                            }
                            table.getRow(i).getCell(j).setText((String) list.get(i).get("strContent"));
                        }
                    }
                }else {
                    //不显示内容
                    if ((boolean) viewConfig.get("showModfiyDate") == true) {
                        String content =  ((String) list.get(i).get("strContent")).replace("ξ", "");
                        if (list.get(i).get("signType") == "2") {
                            //图片
                            String sreg = "/<\\s*img\\s+onerror\\s*=\\s*\\'this.outerHTML\\s*=\\\"(\\S+)\\\"\\'\\s+src\\s*=\\s*[\\'|\\\"](\\S+)[\\'|\\\"]\\s*\\/?>\\s*(.*)/i";
                            Pattern pattern = Pattern.compile(sreg);
                            Matcher matcher = pattern.matcher(content);
                            ArrayList al = new ArrayList();
                            while (matcher.find()) {
                                al.add(matcher.group(0));
                            }
                            if (al.get(1) != null) {
                                try {
                                    XWPFTableRow tableRow = table.createRow();
                                    XWPFParagraph paragraph = tableRow.getCell(j).getParagraphs().get(0);
                                    XWPFRun run = paragraph.getRuns().isEmpty() ? paragraph.createRun() : paragraph.getRuns().get(0);
                                    run.addPicture(null, Document.PICTURE_TYPE_JPEG, "", 100, 200);
                                    //table.openCellRC(i + k + 1, j + 1).Range.InlineShapes.AddPicture(getcontextUrl() + group[2].substring(2, group[2].length));
                                    paragraph.createRun().setText((String) al.get(2));
                                } catch (RuntimeException e) {
                                    table.createRow().getCell(j).setText("" + al.get(1) + al.get(3));
                                }
                            }
                        } else {
                            table.createRow().getCell(j).setText(content);
                        }
                    }
                }
                if (viewConfig.get("showHistoryDate") == "true" && viewConfig.get("showHistoryUserInfo") == "true") {
                    table.createRow().getCell(j).getParagraphs().get(0).setAlignment(disCon ? ParagraphAlignment.RIGHT : ParagraphAlignment.LEFT);
                    table.getRow(i).getCell(j).setText((list.get(i).get("signType") != "2" ? list.get(i).get("userName") : "") + " " + list.get(i).get("subDateTime"));
                    if ((list.get(i).get("signType") == "2")) {
                        try {
                            setCellImg(table.getRow(i).getCell(j), "getcontextUrl() + '/resources/workflow/workform/sign_pic/' + list[i].userId + '.jpg'");
                        } catch (Exception e) {
                            table.getRow(i).getCell(j).setText( list.get(i).get("userName") + " " + list.get(i).get("subDateTime"));
                        }
                    }
                }
            }
        }
    }



    private void insertTable(XWPFTable table, List<Object[]> tableList) throws IOException, InvalidFormatException {
        //创建行,根据需要插入的数据添加新行，不处理表头
        for (int i = 1; i < tableList.size(); i++) {
            table.createRow();
        }
        //遍历表格插入数据
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 1; i < rows.size(); i++) {
            List<XWPFTableCell> cells = rows.get(i).getTableCells();
            for (int j = 0; j < cells.size(); j++) {
                XWPFTableCell cell = cells.get(j);
                setCellLocation(cell, STVerticalJc.CENTER.toString(), STJc.LEFT.toString());
                Object obj = tableList.get(i - 1)[j];
                if(obj instanceof String){
                    cell.setText((String) tableList.get(i - 1)[j]);
                }else if(obj instanceof Picture){
                    Picture pic = (Picture) obj;
                    XWPFParagraph paragraph = cell.getParagraphs().get(0);
                    XWPFRun run = paragraph.getRuns().isEmpty() ? paragraph.createRun() : paragraph.getRuns().get(0);
                    run.addPicture(pic.getPictureData(),pic.getPictureType(), pic.getFileName(), pic.getWidth(), pic.getHeight());;
                }
            }
        }
    }


    /**
     * 设置单元格水平位置和垂直位置
     *
     * @param xwpfTable
     * @param verticalLoction    单元格中内容垂直上TOP，下BOTTOM，居中CENTER，BOTH两端对齐
     * @param horizontalLocation 单元格中内容水平居中center,left居左，right居右，both两端对齐
     */
    public static void setCellLocation(XWPFTable xwpfTable, String verticalLoction, String horizontalLocation) {
        List<XWPFTableRow> rows = xwpfTable.getRows();
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                CTTc cttc = cell.getCTTc();
                CTP ctp = cttc.getPList().get(0);
                CTPPr ctppr = ctp.getPPr();
                if (ctppr == null) {
                    ctppr = ctp.addNewPPr();
                }
                CTJc ctjc = ctppr.getJc();
                if (ctjc == null) {
                    ctjc = ctppr.addNewJc();
                }
                ctjc.setVal(STJc.Enum.forString(horizontalLocation)); //水平居中
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.valueOf(verticalLoction));//垂直居中
            }
        }
    }
    /**
     * 设置单元格水平位置和垂直位置
     *
     * @param cell
     * @param verticalLoction    单元格中内容垂直上TOP，下BOTTOM，居中CENTER，BOTH两端对齐
     * @param horizontalLocation 单元格中内容水平居中center,left居左，right居右，both两端对齐
     */
    public static void setCellLocation(XWPFTableCell cell, String verticalLoction, String horizontalLocation) {
        CTTc cttc = cell.getCTTc();
        CTP ctp = cttc.getPList().get(0);
        CTPPr ctppr = ctp.getPPr();
        if (ctppr == null) {
            ctppr = ctp.addNewPPr();
        }
        CTJc ctjc = ctppr.getJc();
        if (ctjc == null) {
            ctjc = ctppr.addNewJc();
        }
        ctjc.setVal(STJc.Enum.forString(horizontalLocation)); //水平居中
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.valueOf(verticalLoction));//垂直居中
    }
    /**
     * 设置表格位置
     *
     * @param xwpfTable
     * @param location  整个表格居中center,left居左，right居右，both两端对齐
     */
    public static void setTableLocation(XWPFTable xwpfTable, String location) {
        CTTbl cttbl = xwpfTable.getCTTbl();
        CTTblPr tblpr = cttbl.getTblPr() == null ? cttbl.addNewTblPr() : cttbl.getTblPr();
        CTJc cTJc = (CTJc) tblpr.addNewJc();
        cTJc.setVal(STJc.Enum.forString(location));
    }

    /**
     *插入图片
     */
    public void setCellImg(XWPFTableCell cell, String path) throws Exception {
        try{
            //获取单元格的段落
            XWPFParagraph paragraphs = cell.getParagraphs().get(0);
            XWPFRun run = paragraphs.getRuns().isEmpty() ? paragraphs.createRun() : paragraphs.getRuns().get(0);
            int index=0;
            File image = new File(path);
            //判断图片是否存在
            if(!image.exists()){
                return;
            }
            //判断图片的格式
            int format=0;
            if (path.endsWith(".emf")) {
                format = XWPFDocument.PICTURE_TYPE_EMF;
            } else if (path.endsWith(".wmf")) {
                format = XWPFDocument.PICTURE_TYPE_WMF;
            } else if (path.endsWith(".pict")) {
                format = XWPFDocument.PICTURE_TYPE_PICT;
            } else if (path.endsWith(".jpeg") || path.endsWith(".jpg")) {
                format = XWPFDocument.PICTURE_TYPE_JPEG;
            } else if (path.endsWith(".png")) {
                format = XWPFDocument.PICTURE_TYPE_PNG;
            } else if (path.endsWith(".dib")) {
                format = XWPFDocument.PICTURE_TYPE_DIB;
            } else if (path.endsWith(".gif")) {
                format = XWPFDocument.PICTURE_TYPE_GIF;
            } else if (path.endsWith(".tiff")) {
                format = XWPFDocument.PICTURE_TYPE_TIFF;
            } else if (path.endsWith(".eps")) {
                format = XWPFDocument.PICTURE_TYPE_EPS;
            } else if (path.endsWith(".bmp")) {
                format = XWPFDocument.PICTURE_TYPE_BMP;
            } else if (path.endsWith(".wpg")) {
                format = XWPFDocument.PICTURE_TYPE_WPG;
            } else {
                log.error("Unsupported picture: " + path +
                        ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
                return;
            }
            //获取图片文件流
            FileInputStream is = new FileInputStream(path);
            //计算适合文档宽高的图片EMU数值
            BufferedImage read = ImageIO.read(image);
            int width = Units.toEMU(read.getWidth());
            int height = Units.toEMU(read.getHeight());
            //1 EMU = 1/914400英寸= 1/36000 mm,15是word文档中图片能设置的最大宽度cm
            if(width/360000>15){
                NumberFormat f = NumberFormat.getNumberInstance();
                f.setMaximumFractionDigits(0);
                f.setRoundingMode(RoundingMode.UP);
                Double d=width/360000/15d;
                width = Integer.valueOf(f.format(width/d).replace(",",""));
                height = Integer.valueOf(f.format(height/d).replace(",",""));
            }
            run.addPicture(is, format, image.getName(), width, height);
            is.close();
        }catch (Exception e){
            throw e;
        }
    }


}

