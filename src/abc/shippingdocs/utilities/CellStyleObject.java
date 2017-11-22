/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package abc.shippingdocs.utilities;

import abc.shippingdocs.utilities.StyleUtil.Decoration;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 *
 * @author dwink
 */
public class CellStyleObject {

    /**
     * @return the decorations
     */
    public Decoration[] getDecorations() {
        return decorations;
    }

    /**
     * @param decorations the decorations to set
     */
    public void setDecorations(Decoration[] decorations) {
        this.decorations = decorations;
    }

    /**
     * @return the fontName
     */
    public String getFontName() {
        return fontName;
    }

    /**
     * @param fontName the fontName to set
     */
    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    /**
     * @return the fontSize
     */
    public short getFontSize() {
        return fontSize;
    }

    /**
     * @param fontSize the fontSize to set
     */
    public void setFontSize(short fontSize) {
        this.fontSize = fontSize;
    }

    /**
     * @return the alignment
     */
    public HorizontalAlignment getAlignment() {
        return alignment;
    }

    /**
     * @param alignment the alignment to set
     */
    public void setAlignment(HorizontalAlignment alignment) {
        this.alignment = alignment;
    }
    
    private Decoration [] decorations;
    private String fontName;
    private short fontSize;
    private HorizontalAlignment alignment;
    
    
    public CellStyleObject(Decoration [] dec, String name, short size, HorizontalAlignment halign){
        decorations = dec;
        fontName = name;
        fontSize = size;
        alignment = halign;
    }
    
    
}
