/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class TextEffectFormat extends Dispatch {

  public static final String componentName = "Word.TextEffectFormat";

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
   * @return the result is of type Application
   */
  public AppEvents getApplication() {
    return new AppEvents(Dispatch.get(this, "Application").toDispatch());
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

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void toggleVerticalText() {
    Dispatch.call(this, "ToggleVerticalText");
  }

}
