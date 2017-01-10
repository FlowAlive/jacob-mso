/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ThreeDFormat extends Dispatch {

  public static final String componentName = "Word.ThreeDFormat";

  public ThreeDFormat() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public ThreeDFormat(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public ThreeDFormat(String compName) {
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
   * @return the result is of type float
   */
  public float getDepth() {
    return Dispatch.get(this, "Depth").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param depth an input-parameter of type float
   */
  public void setDepth(float depth) {
    Dispatch.put(this, "Depth", new Variant(depth));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type ColorFormat
   */
  public ColorFormat getExtrusionColor() {
    return new ColorFormat(Dispatch.get(this, "ExtrusionColor").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getRotationX() {
    return Dispatch.get(this, "RotationX").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param rotationX an input-parameter of type float
   */
  public void setRotationX(float rotationX) {
    Dispatch.put(this, "RotationX", new Variant(rotationX));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getRotationY() {
    return Dispatch.get(this, "RotationY").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param rotationY an input-parameter of type float
   */
  public void setRotationY(float rotationY) {
    Dispatch.put(this, "RotationY", new Variant(rotationY));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param increment an input-parameter of type float
   */
  public void incrementRotationX(float increment) {
    Dispatch.call(this, "IncrementRotationX", new Variant(increment));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param increment an input-parameter of type float
   */
  public void incrementRotationY(float increment) {
    Dispatch.call(this, "IncrementRotationY", new Variant(increment));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void resetRotation() {
    Dispatch.call(this, "ResetRotation");
  }

}
