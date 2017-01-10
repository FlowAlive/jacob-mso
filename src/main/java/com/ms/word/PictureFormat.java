/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class PictureFormat extends Dispatch {

  public static final String componentName = "Word.PictureFormat";

  public PictureFormat() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public PictureFormat(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public PictureFormat(String compName) {
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
  public float getBrightness() {
    return Dispatch.get(this, "Brightness").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param brightness an input-parameter of type float
   */
  public void setBrightness(float brightness) {
    Dispatch.put(this, "Brightness", new Variant(brightness));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getContrast() {
    return Dispatch.get(this, "Contrast").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param contrast an input-parameter of type float
   */
  public void setContrast(float contrast) {
    Dispatch.put(this, "Contrast", new Variant(contrast));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getCropBottom() {
    return Dispatch.get(this, "CropBottom").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param cropBottom an input-parameter of type float
   */
  public void setCropBottom(float cropBottom) {
    Dispatch.put(this, "CropBottom", new Variant(cropBottom));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getCropLeft() {
    return Dispatch.get(this, "CropLeft").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param cropLeft an input-parameter of type float
   */
  public void setCropLeft(float cropLeft) {
    Dispatch.put(this, "CropLeft", new Variant(cropLeft));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getCropRight() {
    return Dispatch.get(this, "CropRight").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param cropRight an input-parameter of type float
   */
  public void setCropRight(float cropRight) {
    Dispatch.put(this, "CropRight", new Variant(cropRight));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getCropTop() {
    return Dispatch.get(this, "CropTop").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param cropTop an input-parameter of type float
   */
  public void setCropTop(float cropTop) {
    Dispatch.put(this, "CropTop", new Variant(cropTop));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getTransparencyColor() {
    return Dispatch.get(this, "TransparencyColor").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param transparencyColor an input-parameter of type int
   */
  public void setTransparencyColor(int transparencyColor) {
    Dispatch.put(this, "TransparencyColor", new Variant(transparencyColor));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param increment an input-parameter of type float
   */
  public void incrementBrightness(float increment) {
    Dispatch.call(this, "IncrementBrightness", new Variant(increment));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param increment an input-parameter of type float
   */
  public void incrementContrast(float increment) {
    Dispatch.call(this, "IncrementContrast", new Variant(increment));
  }

}
