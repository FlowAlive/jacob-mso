/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class TextEffectFormat extends Dispatch {

  public static final String componentName = "Office.TextEffectFormat";

  public TextEffectFormat() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public TextEffectFormat(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public TextEffectFormat(String compName) {
    super(compName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Object
   */
  public Object getApplication() {
    return Dispatch.get(this, "Application");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getCreator() {
    return Dispatch.get(this, "Creator").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Object
   */
  public Object getParent() {
    return Dispatch.get(this, "Parent");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void toggleVerticalText() {
    Dispatch.call(this, "ToggleVerticalText");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getAlignment() {
    return Dispatch.get(this, "Alignment").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param alignment an input-parameter of type int
   */
  public void setAlignment(int alignment) {
    Dispatch.put(this, "Alignment", new Variant(alignment));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getFontBold() {
    return Dispatch.get(this, "FontBold").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fontBold an input-parameter of type int
   */
  public void setFontBold(int fontBold) {
    Dispatch.put(this, "FontBold", new Variant(fontBold));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getFontItalic() {
    return Dispatch.get(this, "FontItalic").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fontItalic an input-parameter of type int
   */
  public void setFontItalic(int fontItalic) {
    Dispatch.put(this, "FontItalic", new Variant(fontItalic));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getFontName() {
    return Dispatch.get(this, "FontName").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fontName an input-parameter of type String
   */
  public void setFontName(String fontName) {
    Dispatch.put(this, "FontName", fontName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getFontSize() {
    return Dispatch.get(this, "FontSize").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fontSize an input-parameter of type float
   */
  public void setFontSize(float fontSize) {
    Dispatch.put(this, "FontSize", new Variant(fontSize));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getKernedPairs() {
    return Dispatch.get(this, "KernedPairs").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param kernedPairs an input-parameter of type int
   */
  public void setKernedPairs(int kernedPairs) {
    Dispatch.put(this, "KernedPairs", new Variant(kernedPairs));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getNormalizedHeight() {
    return Dispatch.get(this, "NormalizedHeight").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param normalizedHeight an input-parameter of type int
   */
  public void setNormalizedHeight(int normalizedHeight) {
    Dispatch.put(this, "NormalizedHeight", new Variant(normalizedHeight));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getPresetShape() {
    return Dispatch.get(this, "PresetShape").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param presetShape an input-parameter of type int
   */
  public void setPresetShape(int presetShape) {
    Dispatch.put(this, "PresetShape", new Variant(presetShape));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getPresetTextEffect() {
    return Dispatch.get(this, "PresetTextEffect").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param presetTextEffect an input-parameter of type int
   */
  public void setPresetTextEffect(int presetTextEffect) {
    Dispatch.put(this, "PresetTextEffect", new Variant(presetTextEffect));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getRotatedChars() {
    return Dispatch.get(this, "RotatedChars").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param rotatedChars an input-parameter of type int
   */
  public void setRotatedChars(int rotatedChars) {
    Dispatch.put(this, "RotatedChars", new Variant(rotatedChars));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getText() {
    return Dispatch.get(this, "Text").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param text an input-parameter of type String
   */
  public void setText(String text) {
    Dispatch.put(this, "Text", text);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getTracking() {
    return Dispatch.get(this, "Tracking").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param tracking an input-parameter of type float
   */
  public void setTracking(float tracking) {
    Dispatch.put(this, "Tracking", new Variant(tracking));
  }

}
