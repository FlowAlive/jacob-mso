/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.ms.office.MsoEditingType;
import com.ms.office.MsoPresetTextEffect;
import com.ms.office.MsoTextOrientation;
import com.ms.office.MsoTriState;

public class Shapes extends Dispatch {

  public static final String componentName = "Word.Shapes";

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
   * @return the result is of type int
   */
  public int getCount() {
    return Dispatch.get(this, "Count").changeType(Variant.VariantInt).getInt();
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
   * @param index an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape item(Variant index) {
    return new Shape(Dispatch.call(this, "Item", index).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param safeArrayOfPoints an input-parameter of type Variant
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addCurve(Variant safeArrayOfPoints, Variant anchor) {
    return new Shape(Dispatch.call(this, "AddCurve", safeArrayOfPoints, anchor).toDispatch());
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
   * @param orientation an input-parameter of type MsoTextOrientation
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @param height an input-parameter of type float
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addLabel(MsoTextOrientation orientation, float left, float top, float width, float height,
                        Variant anchor) {
    return new Shape(Dispatch.call(this, "AddLabel", orientation, new Variant(left), new Variant(top),
                                   new Variant(width), new Variant(height), anchor).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param orientation an input-parameter of type MsoTextOrientation
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @param height an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addLabel(MsoTextOrientation orientation, float left, float top, float width, float height) {
    return new Shape(Dispatch.call(this, "AddLabel", orientation, new Variant(left), new Variant(top),
                                   new Variant(width), new Variant(height)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param beginX an input-parameter of type float
   * @param beginY an input-parameter of type float
   * @param endX an input-parameter of type float
   * @param endY an input-parameter of type float
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addLine(float beginX, float beginY, float endX, float endY, Variant anchor) {
    return new Shape(Dispatch.call(this, "AddLine", new Variant(beginX), new Variant(beginY), new Variant(endX),
                                   new Variant(endY), anchor).toDispatch());
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
   * @param linkToFile an input-parameter of type Variant
   * @param saveWithDocument an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, Variant linkToFile, Variant saveWithDocument, Variant left, Variant top,
                          Variant width, Variant height, Variant anchor) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, linkToFile, saveWithDocument, left, top, width, height,
                                   anchor).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @param linkToFile an input-parameter of type Variant
   * @param saveWithDocument an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, Variant linkToFile, Variant saveWithDocument, Variant left, Variant top,
                          Variant width, Variant height) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, linkToFile, saveWithDocument, left, top, width, height).
                     toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @param linkToFile an input-parameter of type Variant
   * @param saveWithDocument an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, Variant linkToFile, Variant saveWithDocument, Variant left, Variant top,
                          Variant width) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, linkToFile, saveWithDocument, left, top, width).
                     toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @param linkToFile an input-parameter of type Variant
   * @param saveWithDocument an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, Variant linkToFile, Variant saveWithDocument, Variant left, Variant top) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, linkToFile, saveWithDocument, left, top).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @param linkToFile an input-parameter of type Variant
   * @param saveWithDocument an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, Variant linkToFile, Variant saveWithDocument, Variant left) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, linkToFile, saveWithDocument, left).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @param linkToFile an input-parameter of type Variant
   * @param saveWithDocument an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, Variant linkToFile, Variant saveWithDocument) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, linkToFile, saveWithDocument).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @param linkToFile an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName, Variant linkToFile) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName, linkToFile).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   * @return the result is of type Shape
   */
  public Shape addPicture(String fileName) {
    return new Shape(Dispatch.call(this, "AddPicture", fileName).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param safeArrayOfPoints an input-parameter of type Variant
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addPolyline(Variant safeArrayOfPoints, Variant anchor) {
    return new Shape(Dispatch.call(this, "AddPolyline", safeArrayOfPoints, anchor).toDispatch());
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
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addShape(int type, float left, float top, float width, float height, Variant anchor) {
    return new Shape(Dispatch.call(this, "AddShape", new Variant(type), new Variant(left), new Variant(top),
                                   new Variant(width), new Variant(height), anchor).toDispatch());
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
   * @param presetTextEffect an input-parameter of type MsoPresetTextEffect
   * @param text an input-parameter of type String
   * @param fontName an input-parameter of type String
   * @param fontSize an input-parameter of type float
   * @param fontBold an input-parameter of type MsoTriState
   * @param fontItalic an input-parameter of type MsoTriState
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addTextEffect(MsoPresetTextEffect presetTextEffect, String text, String fontName, float fontSize,
                             MsoTriState fontBold, MsoTriState fontItalic, float left, float top, Variant anchor) {
    return new Shape(Dispatch.callN(this, "AddTextEffect", new Object[] {presetTextEffect, text, fontName,
                                    new Variant(fontSize), fontBold, fontItalic, new Variant(left), new Variant(top),
                                    anchor}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param presetTextEffect an input-parameter of type MsoPresetTextEffect
   * @param text an input-parameter of type String
   * @param fontName an input-parameter of type String
   * @param fontSize an input-parameter of type float
   * @param fontBold an input-parameter of type MsoTriState
   * @param fontItalic an input-parameter of type MsoTriState
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addTextEffect(MsoPresetTextEffect presetTextEffect, String text, String fontName, float fontSize,
                             MsoTriState fontBold, MsoTriState fontItalic, float left, float top) {
    return new Shape(Dispatch.call(this, "AddTextEffect", presetTextEffect, text, fontName, new Variant(fontSize),
                                   fontBold, fontItalic, new Variant(left), new Variant(top)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param orientation an input-parameter of type MsoTextOrientation
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @param height an input-parameter of type float
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addTextbox(MsoTextOrientation orientation, float left, float top, float width, float height,
                          Variant anchor) {
    return new Shape(Dispatch.call(this, "AddTextbox", orientation, new Variant(left), new Variant(top),
                                   new Variant(width), new Variant(height), anchor).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param orientation an input-parameter of type MsoTextOrientation
   * @param left an input-parameter of type float
   * @param top an input-parameter of type float
   * @param width an input-parameter of type float
   * @param height an input-parameter of type float
   * @return the result is of type Shape
   */
  public Shape addTextbox(MsoTextOrientation orientation, float left, float top, float width, float height) {
    return new Shape(Dispatch.call(this, "AddTextbox", orientation, new Variant(left), new Variant(top),
                                   new Variant(width), new Variant(height)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param editingType an input-parameter of type MsoEditingType
   * @param x1 an input-parameter of type float
   * @param y1 an input-parameter of type float
   * @return the result is of type FreeformBuilder
   */
  public FreeformBuilder buildFreeform(MsoEditingType editingType, float x1, float y1) {
    return new FreeformBuilder(Dispatch.call(this, "BuildFreeform", editingType, new Variant(x1), new Variant(y1)).
                               toDispatch());
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
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @param displayAsIcon an input-parameter of type Variant
   * @param iconFileName an input-parameter of type Variant
   * @param iconIndex an input-parameter of type Variant
   * @param iconLabel an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile, Variant displayAsIcon,
                            Variant iconFileName, Variant iconIndex, Variant iconLabel, Variant left, Variant top,
                            Variant width, Variant height, Variant anchor) {
    return new Shape(Dispatch.callN(this, "AddOLEObject", new Object[] {classType, fileName, linkToFile, displayAsIcon,
                                    iconFileName, iconIndex, iconLabel, left, top, width, height, anchor}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @param displayAsIcon an input-parameter of type Variant
   * @param iconFileName an input-parameter of type Variant
   * @param iconIndex an input-parameter of type Variant
   * @param iconLabel an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile, Variant displayAsIcon,
                            Variant iconFileName, Variant iconIndex, Variant iconLabel, Variant left, Variant top,
                            Variant width, Variant height) {
    return new Shape(Dispatch.callN(this, "AddOLEObject", new Object[] {classType, fileName, linkToFile, displayAsIcon,
                                    iconFileName, iconIndex, iconLabel, left, top, width, height}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @param displayAsIcon an input-parameter of type Variant
   * @param iconFileName an input-parameter of type Variant
   * @param iconIndex an input-parameter of type Variant
   * @param iconLabel an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile, Variant displayAsIcon,
                            Variant iconFileName, Variant iconIndex, Variant iconLabel, Variant left, Variant top,
                            Variant width) {
    return new Shape(Dispatch.callN(this, "AddOLEObject", new Object[] {classType, fileName, linkToFile, displayAsIcon,
                                    iconFileName, iconIndex, iconLabel, left, top, width}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @param displayAsIcon an input-parameter of type Variant
   * @param iconFileName an input-parameter of type Variant
   * @param iconIndex an input-parameter of type Variant
   * @param iconLabel an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile, Variant displayAsIcon,
                            Variant iconFileName, Variant iconIndex, Variant iconLabel, Variant left, Variant top) {
    return new Shape(Dispatch.callN(this, "AddOLEObject", new Object[] {classType, fileName, linkToFile, displayAsIcon,
                                    iconFileName, iconIndex, iconLabel, left, top}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @param displayAsIcon an input-parameter of type Variant
   * @param iconFileName an input-parameter of type Variant
   * @param iconIndex an input-parameter of type Variant
   * @param iconLabel an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile, Variant displayAsIcon,
                            Variant iconFileName, Variant iconIndex, Variant iconLabel, Variant left) {
    return new Shape(Dispatch.call(this, "AddOLEObject", classType, fileName, linkToFile, displayAsIcon, iconFileName,
                                   iconIndex, iconLabel, left).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @param displayAsIcon an input-parameter of type Variant
   * @param iconFileName an input-parameter of type Variant
   * @param iconIndex an input-parameter of type Variant
   * @param iconLabel an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile, Variant displayAsIcon,
                            Variant iconFileName, Variant iconIndex, Variant iconLabel) {
    return new Shape(Dispatch.call(this, "AddOLEObject", classType, fileName, linkToFile, displayAsIcon, iconFileName,
                                   iconIndex, iconLabel).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @param displayAsIcon an input-parameter of type Variant
   * @param iconFileName an input-parameter of type Variant
   * @param iconIndex an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile, Variant displayAsIcon,
                            Variant iconFileName, Variant iconIndex) {
    return new Shape(Dispatch.call(this, "AddOLEObject", classType, fileName, linkToFile, displayAsIcon, iconFileName,
                                   iconIndex).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @param displayAsIcon an input-parameter of type Variant
   * @param iconFileName an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile, Variant displayAsIcon,
                            Variant iconFileName) {
    return new Shape(Dispatch.call(this, "AddOLEObject", classType, fileName, linkToFile, displayAsIcon, iconFileName).
                     toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @param displayAsIcon an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile, Variant displayAsIcon) {
    return new Shape(Dispatch.call(this, "AddOLEObject", classType, fileName, linkToFile, displayAsIcon).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @param linkToFile an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName, Variant linkToFile) {
    return new Shape(Dispatch.call(this, "AddOLEObject", classType, fileName, linkToFile).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param fileName an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType, Variant fileName) {
    return new Shape(Dispatch.call(this, "AddOLEObject", classType, fileName).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEObject(Variant classType) {
    return new Shape(Dispatch.call(this, "AddOLEObject", classType).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Shape
   */
  public Shape addOLEObject() {
    return new Shape(Dispatch.call(this, "AddOLEObject").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param anchor an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEControl(Variant classType, Variant left, Variant top, Variant width, Variant height,
                             Variant anchor) {
    return new Shape(Dispatch.call(this, "AddOLEControl", classType, left, top, width, height, anchor).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEControl(Variant classType, Variant left, Variant top, Variant width, Variant height) {
    return new Shape(Dispatch.call(this, "AddOLEControl", classType, left, top, width, height).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEControl(Variant classType, Variant left, Variant top, Variant width) {
    return new Shape(Dispatch.call(this, "AddOLEControl", classType, left, top, width).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEControl(Variant classType, Variant left, Variant top) {
    return new Shape(Dispatch.call(this, "AddOLEControl", classType, left, top).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEControl(Variant classType, Variant left) {
    return new Shape(Dispatch.call(this, "AddOLEControl", classType, left).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param classType an input-parameter of type Variant
   * @return the result is of type Shape
   */
  public Shape addOLEControl(Variant classType) {
    return new Shape(Dispatch.call(this, "AddOLEControl", classType).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Shape
   */
  public Shape addOLEControl() {
    return new Shape(Dispatch.call(this, "AddOLEControl").toDispatch());
  }

}
