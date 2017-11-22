/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package abc.shippingdocs.utilities;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author dwink
 */
public class StyleUtil {

    public enum Decoration {
        BOLD, ITALIC, STRIKETHROUGH, UNDERLINE
    }

    public static Font createFont(Workbook wb, String fontName, Decoration[] decorations, short size) {
        Font font = wb.createFont();
        font.setFontName(fontName);
        for (Decoration d : decorations) {
            switch (d) {
                case BOLD:
                    font.setBold(true);
                    break;

                case ITALIC:
                    font.setItalic(true);
                    break;

                case STRIKETHROUGH:
                    font.setStrikeout(true);
                    break;
            }
        }
        font.setFontHeightInPoints((short) size);
        return font;
    }
    
    public static void addBorderToCell(CellStyle borderStyle, BorderStyle style){        
        borderStyle.setBorderBottom(style);        
        borderStyle.setBorderTop(style);
        borderStyle.setBorderRight(style);
        borderStyle.setBorderLeft(style);
    }

}
