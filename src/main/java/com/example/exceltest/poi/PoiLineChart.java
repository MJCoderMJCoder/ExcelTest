package com.example.exceltest.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.xmlbeans.XmlBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liuzhifeng01
 * @version 1.0
 * @description:
 * @date 2023/9/24 21:36
 */
public class PoiLineChart {
    private static SXSSFWorkbook wb = new SXSSFWorkbook();
    private SXSSFSheet sheet = null;

    public static void main(String[] args) { // 字段名
        // 标题
        List<String> titleArr = new ArrayList<String>();
        //行别   2013年  2014年  2015年  2016年  2017上半年    2017年  2018上半年    2018年  2019上半年    2019年  2020上半年
        titleArr.add("行别");
        titleArr.add("2013年");
        titleArr.add("2014年");
        titleArr.add("2015年");
        titleArr.add("2016年");
        titleArr.add("2017上半年");
        titleArr.add("2017年");
        titleArr.add("2018上半年");
        titleArr.add("2018年");
        titleArr.add("2019上半年");
        titleArr.add("2019年");
        titleArr.add("2020上半年");
        // 模拟数据
        List<Map<String, Object>> dataList = intData();
        String fileName = "模拟数据";
        PoiLineChart demo = new PoiLineChart();
        try {
            // 创建折线图
            demo.createTimeXYChar(titleArr, dataList, fileName);
            //导出到文件
            String savePath = "D:\\JAVA\\poi\\" + fileName + "_" + System.currentTimeMillis() + ".xlsx";
            FileOutputStream out = new FileOutputStream(new File(savePath));
            wb.write(out);
            out.close();
            System.out.println("导出文件到：" + savePath);
//            Runtime.getRuntime().exec("cmd /c start  "+savePath);
//       Runtime.getRuntime().exec("rundll32 url.dll FileProtocolHandler   "+savePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 创建折线图
     *
     * @throws IOException
     */
    public void createTimeXYChar(List<String> titleArr, List<Map<String, Object>> dataList, String fileName) {
        sheet = wb.createSheet(fileName);
        boolean result = drawSheetMap(sheet, dataList, titleArr);
        System.out.println("生成折线图-->" + result);
    }

    /**
     * 生成折线图
     *
     * @param sheet    页签
     * @param dataList 填充数据
     * @param titleArr 图例标题
     * @return
     */
    private boolean drawSheetMap(SXSSFSheet sheet, List<Map<String, Object>> dataList, List<String> titleArr) {
        boolean result = false;
        // 获取sheet名称
        String sheetName = sheet.getSheetName();
        result = drawSheet0Table(sheet, titleArr, dataList);
        // 创建一个画布
        sheet.createDrawingPatriarch();
        XSSFDrawing drawing = sheet.getDrawingPatriarch();
        // 画一个图区域
        // 前四个默认0，开始列 开始行 结束列 结束行
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, dataList.size() + 3, 15, dataList.size() + 25);
        // 创建一个chart对象
        XSSFChart chart = drawing.createChart(anchor);
        CTChart ctChart = chart.getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        //设置画布边框样式
        CTChartSpace space = chart.getCTChartSpace();
        space.addNewRoundedCorners().setVal(false);//去掉圆角

        //设置图表位置
        //CTManualLayout manualLayout = chart.getManualLayout().getCTManualLayout();

        /*
        // 设置图表标题方法2 setting chart title
        ((XSSFChart) chart).setTitleText("图表Demo");
        // 标题的位置（于图表上方或居中覆盖）
        ctChart.getTitle().addNewOverlay().setVal(false);
        ctChart.addNewShowDLblsOverMax().setVal(true);
        CTTitle tt  = ctChart.addNewTitle();
        ctChart.setTitle(tt);
        */
        //设置标题方法2
        CTTitle ctTitle = ctChart.addNewTitle();
        ctTitle.addNewOverlay().setVal(false);// true时与饼图重叠
        ctTitle.addNewTx().addNewRich().addNewBodyPr();
        CTTextBody rich = ctTitle.getTx().getRich();
        rich.addNewLstStyle();
        CTRegularTextRun newR = rich.addNewP().addNewR();
        newR.setT(sheetName);  //标题名称
        newR.addNewRPr().setB(false);
        XmlBoolean xmlBoolean = XmlBoolean.Factory.newInstance();
        xmlBoolean.setBooleanValue(true);
        newR.getRPr().xsetB(xmlBoolean);//是否加粗 0不加粗 1加粗
        newR.getRPr().setLang("zh-CN");
        newR.getRPr().setAltLang("en-US");
        newR.getRPr().setSz(800);//字体大小
        newR.getRPr().addNewLatin().setTypeface("楷体");
        newR.getRPr().getLatin().setCharset((byte) -122);
        newR.getRPr().getLatin().setPitchFamily((byte) 49);
        newR.getRPr().addNewEa().setTypeface("楷体");
        newR.getRPr().getEa().setCharset((byte) -122);
        newR.getRPr().getEa().setPitchFamily((byte) 49);
        // 单元格没有数据时，该点在图表上为0或不显示或跳过直接连接
        ctChart.addNewDispBlanksAs().setVal(STDispBlanksAs.ZERO);
        //是否添加左侧坐标轴
        ctChart.addNewShowDLblsOverMax().setVal(true);
        // 折线图
        CTLineChart ctLineChart = ctPlotArea.addNewLineChart();

        CTBoolean ctBoolean = ctLineChart.addNewVaryColors();
        ctLineChart.addNewGrouping().setVal(STGrouping.STANDARD);

        // 创建序列,并且设置选中区域
        for (int i = 1; i <= dataList.size(); i++) {
            CTLineSer ctLineSer = ctLineChart.addNewSer();

            CTSerTx ctSerTx = ctLineSer.addNewTx();
            // 图例区
            CTStrRef ctStrRef = ctSerTx.addNewStrRef();
            String legendDataRange = new CellRangeAddress(i, i, 0, 0).formatAsString(sheetName, true);
            System.out.println("legendDataRange：" + legendDataRange);
            ctStrRef.setF(legendDataRange);
            ctLineSer.addNewIdx().setVal(i);
            // 横坐标区
            CTAxDataSource cttAxDataSource = ctLineSer.addNewCat();
            ctStrRef = cttAxDataSource.addNewStrRef();
            String axisDataRange = new CellRangeAddress(0, 0, 1, titleArr.size()-1).formatAsString(sheetName, true);
            System.out.println("axisDataRange：" + axisDataRange);
            ctStrRef.setF(axisDataRange);
            // 数据区域
            CTNumDataSource ctNumDataSource = ctLineSer.addNewVal();
            CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
            // 选第1-6行,第1-3列作为数据区域 //1 2 3
            //23456
            String numDataRange = new CellRangeAddress(i, i, 1, titleArr.size()-1).formatAsString(sheetName, true);
            System.out.println("numDataRange：" + numDataRange);
            ctNumRef.setF(numDataRange);
            // 设置标签格式
            ctBoolean.setVal(false);
            CTDLbls newDLbls = ctLineSer.addNewDLbls();
            newDLbls.setShowLegendKey(ctBoolean);
            ctBoolean.setVal(false);
            newDLbls.setShowVal(ctBoolean);
            ctBoolean.setVal(false);
            newDLbls.setShowCatName(ctBoolean);
            newDLbls.setShowSerName(ctBoolean);
            newDLbls.setShowPercent(ctBoolean);
            newDLbls.setShowBubbleSize(ctBoolean);
            newDLbls.setShowLeaderLines(ctBoolean);

            // 是否是平滑曲线
            CTBoolean addNewSmooth = ctLineSer.addNewSmooth();
            addNewSmooth.setVal(true);
            // 是否是堆积曲线?
            CTMarker addNewMarker = ctLineSer.addNewMarker();
            CTMarkerStyle markerStyle = addNewMarker.addNewSymbol();
            //NONE-无 CIRCLE-实心圆圈 STAR-星号  DASH-实线 DIAMOND-方块（菱形） DOT-点 PLUS-加号 SQUARE-正方形 TRIANGLE-三角形 X-X
            markerStyle.setVal(STMarkerStyle.CIRCLE);
            //设置线条颜色
            STHexBinary3 hex = STHexBinary3.Factory.newInstance();
            switch (i) {
                case 1:
                    hex.setStringValue("FE035B");
                    break;
                case 2:
                    hex.setStringValue("0370FE");
                    break;
                case 3:
                    hex.setStringValue("FEDD03");
                    break;
                default:
                    hex.setStringValue("F555FF");
                    break;
            }
            CTShapeProperties ctShapeProperties = ctLineSer.addNewSpPr();
            CTLineProperties lineProperties = ctShapeProperties.addNewLn();
            CTSolidColorFillProperties properties = lineProperties.addNewSolidFill();
            CTSRgbColor rgbColor = properties.addNewSrgbClr();
            rgbColor.xsetVal(hex);
            //设置线条粗细
            lineProperties.setW(20000);
            //设置线条类型 虚线 实线
            CTPresetLineDashProperties ctPresetLineDashProperties = lineProperties.addNewPrstDash();
            ctPresetLineDashProperties.setVal(STPresetLineDashVal.SOLID);

        }


        //告诉条形图它有轴并给它们ID

        int xAxisId = dataList.size() + 1 + 10000;
        int yAxisId = dataList.size() + 2 + 10000;
        ctLineChart.addNewAxId().setVal(xAxisId);
        ctLineChart.addNewAxId().setVal(yAxisId);
        //==========================================
        // 设置x轴属性
        Boolean isXAxisDelete = false;
        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(xAxisId);
        ctCatAx.addNewCrossAx().setVal(yAxisId);
        CTScaling ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);  //X轴排列方向
        ctCatAx.addNewDelete().setVal(isXAxisDelete);// 是否隐藏x轴
        ctCatAx.addNewAxPos().setVal(STAxPos.B);// x轴位置（左右上下）
        ctCatAx.addNewMajorTickMark().setVal(STTickMark.OUT);// 主刻度线在轴上的位置（内、外、交叉、无）
        ctCatAx.addNewMinorTickMark().setVal(STTickMark.NONE);// 次刻度线在轴上的位置（内、外、交叉、无）
        ctCatAx.addNewAuto().setVal(true);
        ctCatAx.addNewLblAlgn().setVal(STLblAlgn.CTR);
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);// 标签（即刻度文字）的位置（轴旁、高、低、无）
        ctCatAx.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new XSSFColor(new Color(58, 76, 78)).getRGB());// x轴颜色,有两种方式设置颜色
//        STHexBinary3 hex_x = STHexBinary3.Factory.newInstance();
//        hex_x.setStringValue("5555FF");
//        ctCatAx.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().xsetVal(hex_x);
//         ctCatAx.addNewCrosses().setVal(STCrosses.MIN);   //X轴以Y轴的最小值或最大值穿过Y轴
        //设置X轴字体
            /*<c:txPr>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
            <a:pPr>
            <a:defRPr baseline="0" sz="800">
            <a:ea charset="-122" pitchFamily="49" typeface="楷体"/>
            </a:defRPr>
            </a:pPr>
            <a:endParaRPr lang="zh-CN"/>
            </a:p>
            </c:txPr>*/
        ctCatAx.addNewTxPr().addNewBodyPr();
        ctCatAx.getTxPr().addNewLstStyle();
        CTTextCharacterProperties properties = ctCatAx.getTxPr().addNewP().addNewPPr().addNewDefRPr();
        properties.setBaseline(0);
        properties.setSz(800);//字体大小
        properties.addNewEa().setTypeface("楷体");
//        properties.getEa().setCharset((byte) -122);
//        properties.getEa().setPitchFamily((byte)49);
        ctCatAx.getTxPr().getPList().get(0).addNewEndParaRPr().setLang("zh-CN");
        //===========================================
        // 设置y轴属性
        CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(yAxisId);
        ctValAx.addNewCrossAx().setVal(xAxisId);
        ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);  //Y轴排列方向
        ctValAx.addNewDelete().setVal(false);// 是否隐藏y轴
        char yAxisPosition = 'L';
        switch (yAxisPosition) {// y轴位置（左右上下）
            case 'L':
                ctValAx.addNewAxPos().setVal(STAxPos.L);
                ctValAx.addNewCrosses().setVal(STCrosses.MIN);// 纵坐标交叉位置（最大、最小、0、指定某一刻度），也可不用设置，此处如果设置为MAX，则上面设置的左侧失效
                break;
            case 'R':
                ctValAx.addNewAxPos().setVal(STAxPos.R);
                ctValAx.addNewCrosses().setVal(STCrosses.MAX);
                break;
            case 'T':
                ctValAx.addNewAxPos().setVal(STAxPos.T);
                break;
            case 'B':
                ctValAx.addNewAxPos().setVal(STAxPos.B);
                break;
            default:
                ctValAx.addNewAxPos().setVal(STAxPos.L);
                ctValAx.addNewCrosses().setVal(STCrosses.MIN);
                break;
        }
        ctValAx.addNewMajorTickMark().setVal(STTickMark.OUT);// 主刻度线在轴上的位置（内、外、交叉、无）
        ctValAx.addNewMinorTickMark().setVal(STTickMark.NONE);// 次刻度线在轴上的位置（内、外、交叉、无）
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);// 标签（即刻度文字）的位置（轴旁、高、低、无）
//        ctValAx.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal( new XSSFColor(new Color(255, 76, 105)).getRGB());// x轴颜色,有两种方式设置颜色
//        STHexBinary3 hex_y = STHexBinary3.Factory.newInstance();
//        hex_y.setStringValue("5555FF");
//        ctValAx.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().xsetVal(hex_y);
//        ctValAx.addNewMajorGridlines().addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new XSSFColor(new Color(134, 134, 134)).getRGB());// 显示主要网格线，并设置颜色

        ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewMin().setVal(0);// 设置纵坐标最小值
        ctScaling.addNewMax().setVal(110);// 设置纵坐标最大值
        //设置纵坐标标题
//        ctValAx.addNewTitle().addNewTx().addNewStrRef()
//        .setF(new CellRangeAddress(0, 0, 0, 0).formatAsString(sheet.getSheetName(), true));

        //图例
       /* <c:legend>
        <c:legendPos val="b"/>
        <c:layout/>
        <c:txPr>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p>
        <a:pPr>
        <a:defRPr baseline="0" sz="800">
        <a:ea charset="-122" pitchFamily="49" typeface="楷体"/>
        </a:defRPr>
        </a:pPr>
        <a:endParaRPr lang="zh-CN"/>
        </a:p>
        </c:txPr>
        </c:legend>*/
        CTLegend legend = ctChart.addNewLegend();
        legend.addNewLegendPos().setVal(STLegendPos.B);//图例位置
        legend.addNewOverlay().setVal(false);
        legend.addNewTxPr().addNewBodyPr();
        legend.getTxPr().addNewLstStyle();
        CTTextCharacterProperties legendProps = legend.getTxPr().addNewP().addNewPPr().addNewDefRPr();
        legendProps.setBaseline(0);
        legendProps.setSz(800);//字体大小
        legendProps.addNewEa().setTypeface("楷体");
        legendProps.getEa().setCharset((byte) -122);
        legendProps.getEa().setPitchFamily((byte) 49);
        legend.getTxPr().getPList().get(0).addNewEndParaRPr().setLang("zh-CN");
        return result;
    }

    /**
     * 生成数据表
     *
     * @param sheet    sheet页对象
     * @param titleArr 表头字段
     * @param dataList 数据
     * @return 是否生成成功
     */
    private boolean drawSheet0Table(SXSSFSheet sheet, List<String> titleArr, List<Map<String, Object>> dataList) {
        // 测试时返回值
        boolean result = true;
        // 初始化表格样式
        List<CellStyle> styleList = tableStyle();
        // 根据数据创建excel第一行标题行
        SXSSFRow row0 = sheet.createRow(0);
        for (int i = 0; i < titleArr.size(); i++) {
            // 设置标题
            row0.createCell(i).setCellValue(titleArr.get(i));
            // 设置标题行样式
            row0.getCell(i).setCellStyle(styleList.get(0));
        }
        // 填充数据2
        for (int i = 0; i < dataList.size(); i++) {
            // 获取每一项的数据
            Map<String, Object> data = dataList.get(i);
            // 设置每一行的字段标题和数据
            SXSSFRow rowi = sheet.createRow(i + 1);
            for (int j = 0; j < data.size(); j++) {
                // 判断是否是标题字段列
                if (j == 0) {
                    rowi.createCell(j).setCellValue((String) data.get("value" + (j + 1)));
                    // 设置左边字段样式
                    sheet.getRow(i + 1).getCell(j).setCellStyle(styleList.get(0));
                } else {
                    rowi.createCell(j).setCellValue(Double.valueOf((String) data.get("value" + (j + 1))));
                    // 设置数据样式
                    sheet.getRow(i + 1).getCell(j).setCellStyle(styleList.get(2));
                }
            }
        }
        return result;
    }

    /**
     * 生成表格样式
     *
     * @return
     */
    private static List<CellStyle> tableStyle() {
        List<CellStyle> cellStyleList = new ArrayList<CellStyle>();
        // 样式准备
        // 标题样式2
        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN); // 下边框
        style.setBorderLeft(BorderStyle.THIN);// 左边框
        style.setBorderTop(BorderStyle.THIN);// 上边框
        style.setBorderRight(BorderStyle.THIN);// 右边框
        style.setAlignment(HorizontalAlignment.CENTER);
        cellStyleList.add(style);
        CellStyle style1 = wb.createCellStyle();
        style1.setBorderBottom(BorderStyle.THIN); // 下边框
        style1.setBorderLeft(BorderStyle.THIN);// 左边框
        style1.setBorderTop(BorderStyle.THIN);// 上边框
        style1.setBorderRight(BorderStyle.THIN);// 右边框
        style1.setAlignment(HorizontalAlignment.CENTER);
        cellStyleList.add(style1);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);// 上边框
        cellStyle.setBorderBottom(BorderStyle.THIN); // 下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);// 左边框
        cellStyle.setBorderRight(BorderStyle.THIN);// 右边框
        cellStyle.setAlignment(HorizontalAlignment.CENTER);// 水平对齐方式
        // cellStyle.setVerticalAlignment(VerticalAlignment.TOP);//垂直对齐方式
        cellStyleList.add(cellStyle);
        return cellStyleList;
    }


    private static List<Map<String, Object>> intData() {
        // 模拟数据
//        工行   21.92% 19.96% 17.10% 15.24% 15.69% 14.35% 15.33% 13.79% 14.41% 13.05% 11.70%    6  6                          工行 -2.71%    富国 -3.23%                            工行 -2.71%    富国 -16.22%
//        建行   21.23% 19.74% 17.27% 15.44% 17.09% 14.80% 16.66% 14.04% 15.62% 13.18% 12.65%    4  2                          建行 -2.97%    汇丰 2.40%                             建行 -2.97%    摩根 -10.00%
//        中行   18.04% 17.28% 14.53% 12.58% 15.20% 12.24% 15.29% 12.06% 14.56% 11.45% 11.10%    11 9                          中行 -3.46%    花旗 3.80%                             中行 -3.46%    汇丰 -8.00%
//        农行   20.89% 19.57% 16.79% 15.14% 16.74% 14.57% 16.72% 13.66% 14.57% 12.43% 11.94%    7  5                          农行 -2.63%    美银 5.67%                             农行 -2.63%    花旗 -6.00%

        List<Map<String, Object>> dataList = new ArrayList<Map<String, Object>>();
        Map<String, Object> dataMap1 = new HashMap<String, Object>();
        dataMap1.put("value1", "工行");
        dataMap1.put("value2", getRandom());
        dataMap1.put("value3", getRandom());
        dataMap1.put("value4", getRandom());
        dataMap1.put("value5", getRandom());
        dataMap1.put("value6", getRandom());
        dataMap1.put("value7", getRandom());
        dataMap1.put("value8", getRandom());
        dataMap1.put("value9", getRandom());
        dataMap1.put("value10", getRandom());
        dataMap1.put("value11", getRandom());
        dataMap1.put("value12", getRandom());

        Map<String, Object> dataMap2 = new HashMap<String, Object>();
        dataMap2.put("value1", "建行");
        dataMap2.put("value2", getRandom());
        dataMap2.put("value3", getRandom());
        dataMap2.put("value4", getRandom());
        dataMap2.put("value5", getRandom());
        dataMap2.put("value6", getRandom());
        dataMap2.put("value7", getRandom());
        dataMap2.put("value8", getRandom());
        dataMap2.put("value9", getRandom());
        dataMap2.put("value10", getRandom());
        dataMap2.put("value11", getRandom());
        dataMap2.put("value12", getRandom());

        Map<String, Object> dataMap3 = new HashMap<String, Object>();
        dataMap3.put("value1", "中行");
        dataMap3.put("value2", getRandom());
        dataMap3.put("value3", getRandom());
        dataMap3.put("value4", getRandom());
        dataMap3.put("value5", getRandom());
        dataMap3.put("value6", getRandom());
        dataMap3.put("value7", getRandom());
        dataMap3.put("value8", getRandom());
        dataMap3.put("value9", getRandom());
        dataMap3.put("value10", getRandom());
        dataMap3.put("value11", getRandom());
        dataMap3.put("value12", getRandom());

        Map<String, Object> dataMap4 = new HashMap<String, Object>();
        dataMap4.put("value1", "农行");
        dataMap4.put("value2", getRandom());
        dataMap4.put("value3", getRandom());
        dataMap4.put("value4", getRandom());
        dataMap4.put("value5", getRandom());
        dataMap4.put("value6", getRandom());
        dataMap4.put("value7", getRandom());
        dataMap4.put("value8", getRandom());
        dataMap4.put("value9", getRandom());
        dataMap4.put("value10", getRandom());
        dataMap4.put("value11", getRandom());
        dataMap4.put("value12", getRandom());

        Map<String, Object> dataMap5 = new HashMap<String, Object>();
        dataMap5.put("value1", "交行");
        dataMap5.put("value2", getRandom());
        dataMap5.put("value3", getRandom());
        dataMap5.put("value4", getRandom());
        dataMap5.put("value5", getRandom());
        dataMap5.put("value6", getRandom());
        dataMap5.put("value7", getRandom());
        dataMap5.put("value8", getRandom());
        dataMap5.put("value9", getRandom());
        dataMap5.put("value10", getRandom());
        dataMap5.put("value11", getRandom());
        dataMap5.put("value12", getRandom());

        Map<String, Object> dataMap6 = new HashMap<String, Object>();
        dataMap6.put("value1", "邮储");
        dataMap6.put("value2", getRandom());
        dataMap6.put("value3", getRandom());
        dataMap6.put("value4", getRandom());
        dataMap6.put("value5", getRandom());
        dataMap6.put("value6", getRandom());
        dataMap6.put("value7", getRandom());
        dataMap6.put("value8", getRandom());
        dataMap6.put("value9", getRandom());
        dataMap6.put("value10", getRandom());
        dataMap6.put("value11", getRandom());
        dataMap6.put("value12", getRandom());

        dataList.add(dataMap1);
        dataList.add(dataMap2);
        dataList.add(dataMap3);
        dataList.add(dataMap4);
        dataList.add(dataMap5);
        dataList.add(dataMap6);

        return dataList;
    }


    public static String getRandom() {
        Double value = Math.random() * 100;
        return String.format("%.2f", value).toString();
    }
}
