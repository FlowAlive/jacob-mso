/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class LineFormat extends Dispatch {

  public static final String componentName = "Office.LineFormat";

  public LineFormat() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public LineFormat(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public LineFormat(String compName) {
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
   * @return the result is of type ColorFormat
   */
  public ColorFormat getBackColor() {
    return new ColorFormat(Dispatch.get(this, "BackColor").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param backColor an input-parameter of type ColorFormat
   */
  public void setBackColor(ColorFormat backColor) {
    Dispatch.put(this, "BackColor", backColor);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBeginArrowheadLength() {
    return Dispatch.get(this, "BeginArrowheadLength").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param beginArrowheadLength an input-parameter of type int
   */
  public void setBeginArrowheadLength(int beginArrowheadLength) {
    Dispatch.put(this, "BeginArrowheadLength", new Variant(beginArrowheadLength));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBeginArrowheadStyle() {
    return Dispatch.get(this, "BeginArrowheadStyle").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param beginArrowheadStyle an input-parameter of type int
   */
  public void setBeginArrowheadStyle(int beginArrowheadStyle) {
    Dispatch.put(this, "BeginArrowheadStyle", new Variant(beginArrowheadStyle));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBeginArrowheadWidth() {
    return Dispatch.get(this, "BeginArrowheadWidth").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param beginArrowheadWidth an input-parameter of type int
   */
  public void setBeginArrowheadWidth(int beginArrowheadWidth) {
    Dispatch.put(this, "BeginArrowheadWidth", new Variant(beginArrowheadWidth));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getDashStyle() {
    return Dispatch.get(this, "DashStyle").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param dashStyle an input-parameter of type int
   */
  public void setDashStyle(int dashStyle) {
    Dispatch.put(this, "DashStyle", new Variant(dashStyle));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getEndArrowheadLength() {
    return Dispatch.get(this, "EndArrowheadLength").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param endArrowheadLength an input-parameter of type int
   */
  public void setEndArrowheadLength(int endArrowheadLength) {
    Dispatch.put(this, "EndArrowheadLength", new Variant(endArrowheadLength));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getEndArrowheadStyle() {
    return Dispatch.get(this, "EndArrowheadStyle").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param endArrowheadStyle an input-parameter of type int
   */
  public void setEndArrowheadStyle(int endArrowheadStyle) {
    Dispatch.put(this, "EndArrowheadStyle", new Variant(endArrowheadStyle));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getEndArrowheadWidth() {
    return Dispatch.get(this, "EndArrowheadWidth").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param endArrowheadWidth an input-parameter of type int
   */
  public void setEndArrowheadWidth(int endArrowheadWidth) {
    Dispatch.put(this, "EndArrowheadWidth", new Variant(endArrowheadWidth));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type ColorFormat
   */
  public ColorFormat getForeColor() {
    return new ColorFormat(Dispatch.get(this, "ForeColor").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param foreColor an input-parameter of type ColorFormat
   */
  public void setForeColor(ColorFormat foreColor) {
    Dispatch.put(this, "ForeColor", foreColor);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getPattern() {
    return Dispatch.get(this, "Pattern").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param pattern an input-parameter of type int
   */
  public void setPattern(int pattern) {
    Dispatch.put(this, "Pattern", new Variant(pattern));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getStyle() {
    return Dispatch.get(this, "Style").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param style an input-parameter of type int
   */
  public void setStyle(int style) {
    Dispatch.put(this, "Style", new Variant(style));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getTransparency() {
    return Dispatch.get(this, "Transparency").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param transparency an input-parameter of type float
   */
  public void setTransparency(float transparency) {
    Dispatch.put(this, "Transparency", new Variant(transparency));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getVisible() {
    return Dispatch.get(this, "Visible").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param visible an input-parameter of type int
   */
  public void setVisible(int visible) {
    Dispatch.put(this, "Visible", new Variant(visible));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getWeight() {
    return Dispatch.get(this, "Weight").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param weight an input-parameter of type float
   */
  public void setWeight(float weight) {
    Dispatch.put(this, "Weight", new Variant(weight));
  }

}
