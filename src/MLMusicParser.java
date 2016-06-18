/**
 * Created by JiayangMiao on 2016/6/17.
 */

import java.io.*;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.jsoup.Jsoup;
import org.jsoup.nodes.*;
import org.jsoup.select.Elements;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.WorkbookFactory;

class MusicInfo {
    String URL;
    String title;
    String artist;
    String composer;

    MusicInfo() {
        URL = "";
        title = "";
        artist = "";
        composer = "";
    }
}

public class MLMusicParser {

    public static void main (String[] args) {
        // Info: the input need to be Starting with div id = "tab-slide-area"
        File inputFile = new File("input.txt");
        try {
            String inputStr = FileUtils.readFileToString(inputFile, "UTF-8");
            Document doc = Jsoup.parseBodyFragment(inputStr);
            // Parse out the div that contains trackc info
            Element trackElement = doc.getElementById("my-unit");
            Elements lis = trackElement.select("div ul li"); // Fetch all the Lis that reside in UL and a div.
            // reason why the input must start with <div id="my-unit">
            int numLis = lis.size();
            ArrayList<MusicInfo> musicList = new ArrayList<>();
            for (int i=0; i < numLis; i++) {
                Element thisElement = lis.get(i);
                MusicInfo thisMusicInfo = new MusicInfo();
                String thisURL = thisElement.attr("data-src").split("\\?")[0];
                thisMusicInfo.URL = thisURL;
                String thisTitle = thisElement.select("div cite").toString();
                thisTitle = thisTitle.replace("<br>", " ");
                thisTitle = thisTitle.replace("<cite>", "");   thisTitle = thisTitle.replace("</cite>", "");
                thisTitle = thisTitle.replace("<small>", "("); thisTitle = thisTitle.replace("</small>", ")");
                thisMusicInfo.title = thisTitle;
                thisMusicInfo.artist = thisElement.select("div span.artist").text();
                thisMusicInfo.composer = thisElement.select("div span.composer").text();
                musicList.add(thisMusicInfo);
            }

            // Parse out the div that contains CD info
            Element cdElement = doc.getElementById("my-deck").getElementsByClass("content").first();
            String cdTitle = cdElement.select("cite").toString();
            cdTitle = cdTitle.replace("<br>", " ");
            cdTitle = cdTitle.replace("<cite>", "");   cdTitle = cdTitle.replace("</cite>", "");
            String cdArtist = cdElement.select("div.artist dl").text();
            String cdCover = cdElement.select("div#audio-room-cd-detail img").attr("src").split("\\?")[0];
            Elements cdDetails = cdElement.select("div#audio-room-cd-detail dl dd");
            String cdRelease = cdDetails.get(0).text();
            String cdLyrics = cdDetails.get(1).text();
            String cdComposer = cdDetails.get(2).text();
            //String cdLyrics = "";
            //String cdComposer = "";
            XSSFWorkbook wb = (XSSFWorkbook) WorkbookFactory.create(new File("MLMusicInfo.xlsx"));
            XSSFSheet sheet = wb.getSheetAt(0);
            int rows = sheet.getPhysicalNumberOfRows();
            XSSFRow row; XSSFCell cell;

            for (int i=0; i<numLis; i++) {
                row = sheet.createRow((short)rows + i);
                cell = row.createCell(0);
                cell.setCellValue(musicList.get(i).title);
                cell = row.createCell(1);
                cell.setCellValue(musicList.get(i).artist);
                cell = row.createCell(2);
                cell.setCellValue(musicList.get(i).composer);
                cell = row.createCell(3);
                cell.setCellValue(musicList.get(i).URL);
            }

            row = sheet.getRow ((short)rows);
                cell = row.createCell(4);
                cell.setCellValue(cdTitle);
                sheet.addMergedRegion(new CellRangeAddress(rows, //first row (0-based)
                        rows+numLis-1, //last row  (0-based)
                        4, //first column (0-based)
                        4  //last column  (0-based)
                ));

                cell = row.createCell(5);
                cell.setCellValue(cdArtist);
                sheet.addMergedRegion(new CellRangeAddress(rows, //first row (0-based)
                        rows+numLis-1, //last row  (0-based)
                        5, //first column (0-based)
                        5  //last column  (0-based)
                ));

                cell = row.createCell(6);
                cell.setCellValue(cdCover);
                sheet.addMergedRegion(new CellRangeAddress(rows, //first row (0-based)
                        rows+numLis-1, //last row  (0-based)
                        6, //first column (0-based)
                        6  //last column  (0-based)
                ));

                cell = row.createCell(7);
                cell.setCellValue(cdRelease);
                sheet.addMergedRegion(new CellRangeAddress(rows, //first row (0-based)
                        rows+numLis-1, //last row  (0-based)
                        7, //first column (0-based)
                        7  //last column  (0-based)
                ));

                cell = row.createCell(8);
                cell.setCellValue(cdLyrics);
                sheet.addMergedRegion(new CellRangeAddress(rows, //first row (0-based)
                        rows+numLis-1, //last row  (0-based)
                        8, //first column (0-based)
                        8  //last column  (0-based)
                ));

                cell = row.createCell(9);
                cell.setCellValue(cdComposer);
                sheet.addMergedRegion(new CellRangeAddress(rows, //first row (0-based)
                        rows+numLis-1, //last row  (0-based)
                        9, //first column (0-based)
                        9  //last column  (0-based)
                ));

            // Style the cell with borders all around.
            CellStyle style = wb.createCellStyle();
            style.setBorderBottom(CellStyle.BORDER_THIN);
            style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            row = sheet.getRow(rows);
            for (int i=4; i<10; i++) {
                cell = row.getCell(i);
                cell.setCellStyle(style);
            }
            row = sheet.getRow(rows+numLis-1);
            for (int i =0; i<4; i++) {
                cell = row.getCell(i);
                cell.setCellStyle(style);
            }

            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("MLMusicInfoTemp.xlsx");
            wb.write(fileOut);
            fileOut.close();
            wb.close();

            File originalFile = new File("MLMusicInfo.xlsx");
            File tempFile = new File("MLMusicInfoTemp.xlsx");
            tempFile.renameTo(originalFile);

            File fileToDelete = FileUtils.getFile("MLMusicInfoTemp.xlsx");
            boolean success = FileUtils.deleteQuietly(fileToDelete);
        }
        catch (Exception e) {
            System.out.println(e.getMessage());
            e.printStackTrace();
        }
    }
}