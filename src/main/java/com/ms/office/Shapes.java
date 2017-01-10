/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Shapes extends Dispatch {

  public static final String componentName = "Office.Shapes";

  public Shapes() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public Shapes(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public Shapes(String compName) {
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
   * @return the result is of type int
   */
  public int getCount() {
    return Dispatch.get(this, "Count").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param index an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape item(Variant index) {
    return new Shape(Dispatch.call(this, "Item", index).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant get_NewEnum() {
    return Dispatch.get(this, "_NewEnum");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type int
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @param height an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addCallout(int type, float left, float top, float width, float height) {
    return new Shape(Dispatch.call(this, "AddCallout", new Variant(type), new Variant(left), new Variant(top),
                                   new Variant(width), new Variant(height)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type int
   * @param beginX an input-parameter of type float
   * @param beginY an input-parameter of type float
   * @param endX an input-parameter of type float
   * @param endY an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addConnector(int type, float beginX, float beginY, float endX, float endY) {
    return new Shape(Dispatch.call(this, "AddConnector", new Variant(type), new Variant(beginX), new Variant(beginY),
                                   new Variant(endX), new Variant(endY)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param safeArrayOfPoints an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addCurve(Variant safeArrayOfPoints) {
    return new Shape(Dispatch.call(this, "AddCurve", safeArrayOfPoints).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param orientation an input-parameter of type int
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @param height an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addLabel(int orientation, float left, float top, float width, float height) {
    return new Shape(Dispatch.call(this, "AddLabel", new Variant(orientation), new Variant(left), new Variant(top),
                                   new Variant(width), new Variant(height)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param beginX an input-parameter of type float
   * @param beginY an input-parameter of type float
   * @param endX an input-parameter of type float
   * @param endY an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addLine(float beginX, float beginY, float endX, float endY) {
    return new Shape(Dispatch.call(this, "AddLine", new Variant(beginX), new Variant(beginY), new Variant(endX),
                                   new Variant(endY)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @param linkToFile an input-parameter of type int
   * @param saveWithDocument an input-parameter of type int
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @param height an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, int linkToFile, int saveWithDocument, float left, float top, float width,
                          float height) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, new Variant(linkToFile), new Variant(saveWithDocument),
                                   new Variant(left), new Variant(top), new Variant(width), new Variant(height)).
                     toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @param linkToFile an input-parameter of type int
   * @param saveWithDocument an input-parameter of type int
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, int linkToFile, int saveWithDocument, float left, float top, float width) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, new Variant(linkToFile), new Variant(saveWithDocument),
                                   new Variant(left), new Variant(top), new Variant(width)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @param linkToFile an input-parameter of type int
   * @param saveWithDocument an input-parameter of type int
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, int linkToFile, int saveWithDocument, float left, float top) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, new Variant(linkToFile), new Variant(saveWithDocument),
                                   new Variant(left), new Variant(top)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param safeArrayOfPoints an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addPolyline(Variant safeArrayOfPoints) {
    return new Shape(Dispatch.call(this, "AddPolyline", safeArrayOfPoints).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type int
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @param height an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addShape(int type, float left, float top, float width, float height) {
    return new Shape(Dispatch.call(this, "AddShape", new Variant(type), new Variant(left), new Variant(top),
                                   new Variant(width), new Variant(height)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param presetTextEffect an input-parameter of type int
   * @param text an input-parameter of type String
   * @param fontName an input-parameter of type String
   * @param fontSize an input-parameter of type float
   * @param fontBold an input-parameter of type int
   * @param fontItalic an input-parameter of type int
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addTextEffect(int presetTextEffect, String text, String fontName, float fontSize, int fontBold,
                             int fontItalic, float left, float top) {
    return new Shape(Dispatch.call(this, "AddTextEffect", new Variant(presetTextEffect), text, fontName,
                                   new Variant(fontSize), new Variant(fontBold), new Variant(fontItalic),
                                   new Variant(left), new Variant(top)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param orientation an input-parameter of type int
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @param height an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addTextbox(int orientation, float left, float top, float width, float height) {
    return new Shape(Dispatch.call(this, "AddTextbox", new Variant(orientation), new Variant(left), new Variant(top),
                                   new Variant(width), new Variant(height)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param editingType an input-parameter of type int
   * @param x1 an input-parameter of type float
   * @param y1 an input-parameter of type float
   * @return the result is of type FreeformBuilder
   */
  public FreeformBuilder buildFreeform(int editingType, float x1, float y1) {
    return new FreeformBuilder(Dispatch.call(this, "BuildFreeform", new Variant(editingType), new Variant(x1),
                                             new Variant(y1)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param index an input-parameter of type Variant
   * @return the result is of type ShapeRange
   */
  public ShapeRange range(Variant index) {
    return new ShapeRange(Dispatch.call(this, "Range", index).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void selectAll() {
    Dispatch.call(this, "SelectAll");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Shape
   */
  public Shape getBackground() {
    return new Shape(Dispatch.get(this, "Background").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Shape
   */
  public Shape getDefault() {
    return new Shape(Dispatch.get(this, "Default").toDispatch());
  }

}
