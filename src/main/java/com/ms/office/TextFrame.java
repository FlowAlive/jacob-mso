/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class TextFrame extends Dispatch {

  public static final String componentName = "Office.TextFrame";

  public TextFrame() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public TextFrame(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public TextFrame(String compName) {
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
   * @return the result is of type float
   */
  public float getMarginBottom() {
    return Dispatch.get(this, "MarginBottom").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param marginBottom an input-parameter of type float
   */
  public void setMarginBottom(float marginBottom) {
    Dispatch.put(this, "MarginBottom", new Variant(marginBottom));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getMarginLeft() {
    return Dispatch.get(this, "MarginLeft").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param marginLeft an input-parameter of type float
   */
  public void setMarginLeft(float marginLeft) {
    Dispatch.put(this, "MarginLeft", new Variant(marginLeft));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getMarginRight() {
    return Dispatch.get(this, "MarginRight").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param marginRight an input-parameter of type float
   */
  public void setMarginRight(float marginRight) {
    Dispatch.put(this, "MarginRight", new Variant(marginRight));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getMarginTop() {
    return Dispatch.get(this, "MarginTop").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param marginTop an input-parameter of type float
   */
  public void setMarginTop(float marginTop) {
    Dispatch.put(this, "MarginTop", new Variant(marginTop));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getOrientation() {
    return Dispatch.get(this, "Orientation").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param orientation an input-parameter of type int
   */
  public void setOrientation(int orientation) {
    Dispatch.put(this, "Orientation", new Variant(orientation));
  }

}
