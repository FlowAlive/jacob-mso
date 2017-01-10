/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class WebPageFont extends Dispatch {

  public static final String componentName = "Office.WebPageFont";

  public WebPageFont() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public WebPageFont(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public WebPageFont(String compName) {
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
   * @return the result is of type String
   */
  public String getProportionalFont() {
    return Dispatch.get(this, "ProportionalFont").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param proportionalFont an input-parameter of type String
   */
  public void setProportionalFont(String proportionalFont) {
    Dispatch.put(this, "ProportionalFont", proportionalFont);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getProportionalFontSize() {
    return Dispatch.get(this, "ProportionalFontSize").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param proportionalFontSize an input-parameter of type float
   */
  public void setProportionalFontSize(float proportionalFontSize) {
    Dispatch.put(this, "ProportionalFontSize", new Variant(proportionalFontSize));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getFixedWidthFont() {
    return Dispatch.get(this, "FixedWidthFont").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fixedWidthFont an input-parameter of type String
   */
  public void setFixedWidthFont(String fixedWidthFont) {
    Dispatch.put(this, "FixedWidthFont", fixedWidthFont);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getFixedWidthFontSize() {
    return Dispatch.get(this, "FixedWidthFontSize").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fixedWidthFontSize an input-parameter of type float
   */
  public void setFixedWidthFontSize(float fixedWidthFontSize) {
    Dispatch.put(this, "FixedWidthFontSize", new Variant(fixedWidthFontSize));
  }

}
