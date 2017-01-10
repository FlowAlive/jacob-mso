/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ShadowFormat extends Dispatch {

  public static final String componentName = "Word.ShadowFormat";

  public ShadowFormat() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public ShadowFormat(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public ShadowFormat(String compName) {
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
   * @return the result is of type ColorFormat
   */
  public ColorFormat getForeColor() {
    return new ColorFormat(Dispatch.get(this, "ForeColor").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getOffsetX() {
    return Dispatch.get(this, "OffsetX").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param offsetX an input-parameter of type float
   */
  public void setOffsetX(float offsetX) {
    Dispatch.put(this, "OffsetX", new Variant(offsetX));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getOffsetY() {
    return Dispatch.get(this, "OffsetY").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param offsetY an input-parameter of type float
   */
  public void setOffsetY(float offsetY) {
    Dispatch.put(this, "OffsetY", new Variant(offsetY));
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
   * @param increment an input-parameter of type float
   */
  public void incrementOffsetX(float increment) {
    Dispatch.call(this, "IncrementOffsetX", new Variant(increment));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param increment an input-parameter of type float
   */
  public void incrementOffsetY(float increment) {
    Dispatch.call(this, "IncrementOffsetY", new Variant(increment));
  }

}
