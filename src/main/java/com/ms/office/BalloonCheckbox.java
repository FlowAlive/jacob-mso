/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class BalloonCheckbox extends Dispatch {

  public static final String componentName = "Office.BalloonCheckbox";

  public BalloonCheckbox() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public BalloonCheckbox(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public BalloonCheckbox(String compName) {
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
  public String getItem() {
    return Dispatch.get(this, "Item").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getName() {
    return Dispatch.get(this, "Name").toString();
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
   * @param checked an input-parameter of type boolean
   */
  public void setChecked(boolean checked) {
    Dispatch.put(this, "Checked", new Variant(checked));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getChecked() {
    return Dispatch.get(this, "Checked").changeType(Variant.VariantBoolean).getBoolean();
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
   * @return the result is of type String
   */
  public String getText() {
    return Dispatch.get(this, "Text").toString();
  }

}
