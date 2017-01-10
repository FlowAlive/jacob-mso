/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.ms.office.MsoFlipCmd;
import com.ms.office.MsoScaleFrom;
import com.ms.office.MsoTriState;
import com.ms.office.MsoZOrderCmd;
import com.ms.office.Script;

public class Shape extends Dispatch {

  public static final String componentName = "Word.Shape";

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
   * @return the result is of type Adjustments
   */
  public Adjustments getAdjustments() {
    return new Adjustments(Dispatch.get(this, "Adjustments").toDispatch());
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
   * @param lockAspectRatio an input-parameter of type MsoTriState
   */
  public void setLockAspectRatio(MsoTriState lockAspectRatio) {
    Dispatch.put(this, "LockAspectRatio", lockAspectRatio);
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
   * @return the result is of type Variant
   */
  public Variant getVertices() {
    return Dispatch.get(this, "Vertices");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param visible an input-parameter of type MsoTriState
   */
  public void setVisible(MsoTriState visible) {
    Dispatch.put(this, "Visible", visible);
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
   * @return the result is of type Hyperlink
   */
  public Hyperlink getHyperlink() {
    return new Hyperlink(Dispatch.get(this, "Hyperlink").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getRelativeHorizontalPosition() {
    return Dispatch.get(this, "RelativeHorizontalPosition").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param relativeHorizontalPosition an input-parameter of type int
   */
  public void setRelativeHorizontalPosition(int relativeHorizontalPosition) {
    Dispatch.put(this, "RelativeHorizontalPosition", new Variant(relativeHorizontalPosition));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getRelativeVerticalPosition() {
    return Dispatch.get(this, "RelativeVerticalPosition").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param relativeVerticalPosition an input-parameter of type int
   */
  public void setRelativeVerticalPosition(int relativeVerticalPosition) {
    Dispatch.put(this, "RelativeVerticalPosition", new Variant(relativeVerticalPosition));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getLockAnchor() {
    return Dispatch.get(this, "LockAnchor").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param lockAnchor an input-parameter of type int
   */
  public void setLockAnchor(int lockAnchor) {
    Dispatch.put(this, "LockAnchor", new Variant(lockAnchor));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type WrapFormat
   */
  public WrapFormat getWrapFormat() {
    return new WrapFormat(Dispatch.get(this, "WrapFormat").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type OLEFormat
   */
  public OLEFormat getOLEFormat() {
    return new OLEFormat(Dispatch.get(this, "OLEFormat").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Range
   */
  public Range getAnchor() {
    return new Range(Dispatch.get(this, "Anchor").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type LinkFormat
   */
  public LinkFormat getLinkFormat() {
    return new LinkFormat(Dispatch.get(this, "LinkFormat").toDispatch());
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
   * @param flipCmd an input-parameter of type MsoFlipCmd
   */
  public void flip(MsoFlipCmd flipCmd) {
    Dispatch.call(this, "Flip", flipCmd);
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
   * @param relativeToOriginalSize an input-parameter of type MsoTriState
   * @param scale an input-parameter of type MsoScaleFrom
   */
  public void scaleHeight(float factor, MsoTriState relativeToOriginalSize, MsoScaleFrom scale) {
    Dispatch.call(this, "ScaleHeight", new Variant(factor), relativeToOriginalSize, scale);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param factor an input-parameter of type float
   * @param relativeToOriginalSize an input-parameter of type MsoTriState
   */
  public void scaleHeight(float factor, MsoTriState relativeToOriginalSize) {
    Dispatch.call(this, "ScaleHeight", new Variant(factor), relativeToOriginalSize);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param factor an input-parameter of type float
   * @param relativeToOriginalSize an input-parameter of type MsoTriState
   * @param scale an input-parameter of type MsoScaleFrom
   */
  public void scaleWidth(float factor, MsoTriState relativeToOriginalSize, MsoScaleFrom scale) {
    Dispatch.call(this, "ScaleWidth", new Variant(factor), relativeToOriginalSize, scale);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param factor an input-parameter of type float
   * @param relativeToOriginalSize an input-parameter of type MsoTriState
   */
  public void scaleWidth(float factor, MsoTriState relativeToOriginalSize) {
    Dispatch.call(this, "ScaleWidth", new Variant(factor), relativeToOriginalSize);
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
   * @param zOrderCmd an input-parameter of type MsoZOrderCmd
   */
  public void zOrder(MsoZOrderCmd zOrderCmd) {
    Dispatch.call(this, "ZOrder", zOrderCmd);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type InlineShape
   */
  public InlineShape convertToInlineShape() {
    return new InlineShape(Dispatch.call(this, "ConvertToInlineShape").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Frame
   */
  public Frame convertToFrame() {
    return new Frame(Dispatch.call(this, "ConvertToFrame").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void activate() {
    Dispatch.call(this, "Activate");
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

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Script
   */
  public Script getScript() {
    return new Script(Dispatch.get(this, "Script").toDispatch());
  }

}
