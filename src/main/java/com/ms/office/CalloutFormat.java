/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class CalloutFormat extends Dispatch {

  public static final String componentName = "Office.CalloutFormat";

  public CalloutFormat() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public CalloutFormat(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public CalloutFormat(String compName) {
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
  public void automaticLength() {
    Dispatch.call(this, "AutomaticLength");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param drop an input-parameter of type float
   */
  public void customDrop(float drop) {
    Dispatch.call(this, "CustomDrop", new Variant(drop));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param length an input-parameter of type float
   */
  public void customLength(float length) {
    Dispatch.call(this, "CustomLength", new Variant(length));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param dropType an input-parameter of type int
   */
  public void presetDrop(int dropType) {
    Dispatch.call(this, "PresetDrop", new Variant(dropType));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getAccent() {
    return Dispatch.get(this, "Accent").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param accent an input-parameter of type int
   */
  public void setAccent(int accent) {
    Dispatch.put(this, "Accent", new Variant(accent));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getAngle() {
    return Dispatch.get(this, "Angle").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param angle an input-parameter of type int
   */
  public void setAngle(int angle) {
    Dispatch.put(this, "Angle", new Variant(angle));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getAutoAttach() {
    return Dispatch.get(this, "AutoAttach").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param autoAttach an input-parameter of type int
   */
  public void setAutoAttach(int autoAttach) {
    Dispatch.put(this, "AutoAttach", new Variant(autoAttach));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getAutoLength() {
    return Dispatch.get(this, "AutoLength").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBorder() {
    return Dispatch.get(this, "Border").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param border an input-parameter of type int
   */
  public void setBorder(int border) {
    Dispatch.put(this, "Border", new Variant(border));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getDrop() {
    return Dispatch.get(this, "Drop").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getDropType() {
    return Dispatch.get(this, "DropType").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getGap() {
    return Dispatch.get(this, "Gap").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param gap an input-parameter of type float
   */
  public void setGap(float gap) {
    Dispatch.put(this, "Gap", new Variant(gap));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getLength() {
    return Dispatch.get(this, "Length").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getType() {
    return Dispatch.get(this, "Type").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type int
   */
  public void setType(int type) {
    Dispatch.put(this, "Type", new Variant(type));
  }

}
