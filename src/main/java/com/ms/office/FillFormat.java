/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class FillFormat extends Dispatch {

  public static final String componentName = "Office.FillFormat";

  public FillFormat() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public FillFormat(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public FillFormat(String compName) {
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
  public void background() {
    Dispatch.call(this, "Background");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param style an input-parameter of type int
   * @param variant an input-parameter of type int
   * @param degree an input-parameter of type float
   */
  public void oneColorGradient(int style, int variant, float degree) {
    Dispatch.call(this, "OneColorGradient", new Variant(style), new Variant(variant), new Variant(degree));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param pattern an input-parameter of type int
   */
  public void patterned(int pattern) {
    Dispatch.call(this, "Patterned", new Variant(pattern));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param style an input-parameter of type int
   * @param variant an input-parameter of type int
   * @param presetGradientType an input-parameter of type int
   */
  public void presetGradient(int style, int variant, int presetGradientType) {
    Dispatch.call(this, "PresetGradient", new Variant(style), new Variant(variant), new Variant(presetGradientType));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param presetTexture an input-parameter of type int
   */
  public void presetTextured(int presetTexture) {
    Dispatch.call(this, "PresetTextured", new Variant(presetTexture));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void solid() {
    Dispatch.call(this, "Solid");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param style an input-parameter of type int
   * @param variant an input-parameter of type int
   */
  public void twoColorGradient(int style, int variant) {
    Dispatch.call(this, "TwoColorGradient", new Variant(style), new Variant(variant));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param pictureFile an input-parameter of type String
   */
  public void userPicture(String pictureFile) {
    Dispatch.call(this, "UserPicture", pictureFile);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param textureFile an input-parameter of type String
   */
  public void userTextured(String textureFile) {
    Dispatch.call(this, "UserTextured", textureFile);
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
  public int getGradientColorType() {
    return Dispatch.get(this, "GradientColorType").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getGradientDegree() {
    return Dispatch.get(this, "GradientDegree").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getGradientStyle() {
    return Dispatch.get(this, "GradientStyle").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getGradientVariant() {
    return Dispatch.get(this, "GradientVariant").changeType(Variant.VariantInt).getInt();
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
   * @return the result is of type int
   */
  public int getPresetGradientType() {
    return Dispatch.get(this, "PresetGradientType").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getPresetTexture() {
    return Dispatch.get(this, "PresetTexture").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getTextureName() {
    return Dispatch.get(this, "TextureName").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getTextureType() {
    return Dispatch.get(this, "TextureType").changeType(Variant.VariantInt).getInt();
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
  public int getType() {
    return Dispatch.get(this, "Type").changeType(Variant.VariantInt).getInt();
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

}
