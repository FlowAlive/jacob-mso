/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Balloon extends Dispatch {

  public static final String componentName = "Office.Balloon";

  public Balloon() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public Balloon(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public Balloon(String compName) {
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
   * @return the result is of type Object
   */
  public Object getCheckboxes() {
    return Dispatch.get(this, "Checkboxes");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Object
   */
  public Object getLabels() {
    return Dispatch.get(this, "Labels");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param balloonType an input-parameter of type int
   */
  public void setBalloonType(int balloonType) {
    Dispatch.put(this, "BalloonType", new Variant(balloonType));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBalloonType() {
    return Dispatch.get(this, "BalloonType").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param icon an input-parameter of type int
   */
  public void setIcon(int icon) {
    Dispatch.put(this, "Icon", new Variant(icon));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getIcon() {
    return Dispatch.get(this, "Icon").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param heading an input-parameter of type String
   */
  public void setHeading(String heading) {
    Dispatch.put(this, "Heading", heading);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getHeading() {
    return Dispatch.get(this, "Heading").toString();
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

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param mode an input-parameter of type int
   */
  public void setMode(int mode) {
    Dispatch.put(this, "Mode", new Variant(mode));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getMode() {
    return Dispatch.get(this, "Mode").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param animation an input-parameter of type int
   */
  public void setAnimation(int animation) {
    Dispatch.put(this, "Animation", new Variant(animation));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getAnimation() {
    return Dispatch.get(this, "Animation").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param button an input-parameter of type int
   */
  public void setButton(int button) {
    Dispatch.put(this, "Button", new Variant(button));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getButton() {
    return Dispatch.get(this, "Button").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param callback an input-parameter of type String
   */
  public void setCallback(String callback) {
    Dispatch.put(this, "Callback", callback);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getCallback() {
    return Dispatch.get(this, "Callback").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param private an input-parameter of type int
   */
  public void setPrivate(int _private) {
    Dispatch.put(this, "Private", new Variant(_private));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getPrivate() {
    return Dispatch.get(this, "Private").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param left an input-parameter of type int
   * @param top an input-parameter of type int
   * @param right an input-parameter of type int
   * @param bottom an input-parameter of type int
   */
  public void setAvoidRectangle(int left, int top, int right, int bottom) {
    Dispatch.call(this, "SetAvoidRectangle", new Variant(left), new Variant(top), new Variant(right),
                  new Variant(bottom));
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
   * @return the result is of type int
   */
  public int show() {
    return Dispatch.call(this, "Show").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void close() {
    Dispatch.call(this, "Close");
  }

}
