/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Shape extends Dispatch {

  public static final String componentName = "Office.Shape";

  public Shape() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public Shape(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public Shape(String compName) {
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
  public void apply() {
    Dispatch.call(this, "Apply");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void delete() {
    Dispatch.call(this, "Delete");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Shape
   */
  public Shape duplicate() {
    return new Shape(Dispatch.call(this, "Duplicate").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param flipCmd an input-parameter of type int
   */
  public void flip(int flipCmd) {
    Dispatch.call(this, "Flip", new Variant(flipCmd));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param increment an input-parameter of type float
   */
  public void incrementLeft(float increment) {
    Dispatch.call(this, "IncrementLeft", new Variant(increment));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param increment an input-parameter of type float
   */
  public void incrementRotation(float increment) {
    Dispatch.call(this, "IncrementRotation", new Variant(increment));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param increment an input-parameter of type float
   */
  public void incrementTop(float increment) {
    Dispatch.call(this, "IncrementTop", new Variant(increment));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void pickUp() {
    Dispatch.call(this, "PickUp");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void rerouteConnections() {
    Dispatch.call(this, "RerouteConnections");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param factor an input-parameter of type float
   * @param relativeToOriginalSize an input-parameter of type int
   * @param fScale an input-parameter of type int
   */
  public void scaleHeight(float factor, int relativeToOriginalSize, int fScale) {
    Dispatch.call(this, "ScaleHeight", new Variant(factor), new Variant(relativeToOriginalSize), new Variant(fScale));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param factor an input-parameter of type float
   * @param relativeToOriginalSize an input-parameter of type int
   */
  public void scaleHeight(float factor, int relativeToOriginalSize) {
    Dispatch.call(this, "ScaleHeight", new Variant(factor), new Variant(relativeToOriginalSize));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param factor an input-parameter of type float
   * @param relativeToOriginalSize an input-parameter of type int
   * @param fScale an input-parameter of type int
   */
  public void scaleWidth(float factor, int relativeToOriginalSize, int fScale) {
    Dispatch.call(this, "ScaleWidth", new Variant(factor), new Variant(relativeToOriginalSize), new Variant(fScale));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param factor an input-parameter of type float
   * @param relativeToOriginalSize an input-parameter of type int
   */
  public void scaleWidth(float factor, int relativeToOriginalSize) {
    Dispatch.call(this, "ScaleWidth", new Variant(factor), new Variant(relativeToOriginalSize));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param replace an input-parameter of type Variant
   */
  public void select(Variant replace) {
    Dispatch.call(this, "Select", replace);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void select() {
    Dispatch.call(this, "Select");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void setShapesDefaultProperties() {
    Dispatch.call(this, "SetShapesDefaultProperties");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type ShapeRange
   */
  public ShapeRange ungroup() {
    return new ShapeRange(Dispatch.call(this, "Ungroup").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param zOrderCmd an input-parameter of type int
   */
  public void zOrder(int zOrderCmd) {
    Dispatch.call(this, "ZOrder", new Variant(zOrderCmd));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Adjustments
   */
  public Adjustments getAdjustments() {
    return new Adjustments(Dispatch.get(this, "Adjustments").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getAutoShapeType() {
    return Dispatch.get(this, "AutoShapeType").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param autoShapeType an input-parameter of type int
   */
  public void setAutoShapeType(int autoShapeType) {
    Dispatch.put(this, "AutoShapeType", new Variant(autoShapeType));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBlackWhiteMode() {
    return Dispatch.get(this, "BlackWhiteMode").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param blackWhiteMode an input-parameter of type int
   */
  public void setBlackWhiteMode(int blackWhiteMode) {
    Dispatch.put(this, "BlackWhiteMode", new Variant(blackWhiteMode));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type CalloutFormat
   */
  public CalloutFormat getCallout() {
    return new CalloutFormat(Dispatch.get(this, "Callout").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getConnectionSiteCount() {
    return Dispatch.get(this, "ConnectionSiteCount").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getConnector() {
    return Dispatch.get(this, "Connector").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type ConnectorFormat
   */
  public ConnectorFormat getConnectorFormat() {
    return new ConnectorFormat(Dispatch.get(this, "ConnectorFormat").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type FillFormat
   */
  public FillFormat getFill() {
    return new FillFormat(Dispatch.get(this, "Fill").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type GroupShapes
   */
  public GroupShapes getGroupItems() {
    return new GroupShapes(Dispatch.get(this, "GroupItems").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getHeight() {
    return Dispatch.get(this, "Height").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param height an input-parameter of type float
   */
  public void setHeight(float height) {
    Dispatch.put(this, "Height", new Variant(height));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getHorizontalFlip() {
    return Dispatch.get(this, "HorizontalFlip").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getLeft() {
    return Dispatch.get(this, "Left").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param left an input-parameter of type float
   */
  public void setLeft(float left) {
    Dispatch.put(this, "Left", new Variant(left));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type LineFormat
   */
  public LineFormat getLine() {
    return new LineFormat(Dispatch.get(this, "Line").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getLockAspectRatio() {
    return Dispatch.get(this, "LockAspectRatio").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param lockAspectRatio an input-parameter of type int
   */
  public void setLockAspectRatio(int lockAspectRatio) {
    Dispatch.put(this, "LockAspectRatio", new Variant(lockAspectRatio));
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
   * @param name an input-parameter of type String
   */
  public void setName(String name) {
    Dispatch.put(this, "Name", name);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type ShapeNodes
   */
  public ShapeNodes getNodes() {
    return new ShapeNodes(Dispatch.get(this, "Nodes").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getRotation() {
    return Dispatch.get(this, "Rotation").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param rotation an input-parameter of type float
   */
  public void setRotation(float rotation) {
    Dispatch.put(this, "Rotation", new Variant(rotation));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type PictureFormat
   */
  public PictureFormat getPictureFormat() {
    return new PictureFormat(Dispatch.get(this, "PictureFormat").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type ShadowFormat
   */
  public ShadowFormat getShadow() {
    return new ShadowFormat(Dispatch.get(this, "Shadow").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type TextEffectFormat
   */
  public TextEffectFormat getTextEffect() {
    return new TextEffectFormat(Dispatch.get(this, "TextEffect").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type TextFrame
   */
  public TextFrame getTextFrame() {
    return new TextFrame(Dispatch.get(this, "TextFrame").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type ThreeDFormat
   */
  public ThreeDFormat getThreeD() {
    return new ThreeDFormat(Dispatch.get(this, "ThreeD").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getTop() {
    return Dispatch.get(this, "Top").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param top an input-parameter of type float
   */
  public void setTop(float top) {
    Dispatch.put(this, "Top", new Variant(top));
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
  public int getVerticalFlip() {
    return Dispatch.get(this, "VerticalFlip").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant getVertices() {
    return Dispatch.get(this, "Vertices");
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

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getWidth() {
    return Dispatch.get(this, "Width").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param width an input-parameter of type float
   */
  public void setWidth(float width) {
    Dispatch.put(this, "Width", new Variant(width));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getZOrderPosition() {
    return Dispatch.get(this, "ZOrderPosition").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Script
   */
  public Script getScript() {
    return new Script(Dispatch.get(this, "Script").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getAlternativeText() {
    return Dispatch.get(this, "AlternativeText").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param alternativeText an input-parameter of type String
   */
  public void setAlternativeText(String alternativeText) {
    Dispatch.put(this, "AlternativeText", alternativeText);
  }

}
