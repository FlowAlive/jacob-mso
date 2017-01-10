/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ConnectorFormat extends Dispatch {

  public static final String componentName = "Office.ConnectorFormat";

  public ConnectorFormat() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public ConnectorFormat(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public ConnectorFormat(String compName) {
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
   * @param connectedShape an input-parameter of type Shape
   * @param connectionSite an input-parameter of type int
   */
  public void beginConnect(Shape connectedShape, int connectionSite) {
    Dispatch.call(this, "BeginConnect", connectedShape, new Variant(connectionSite));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void beginDisconnect() {
    Dispatch.call(this, "BeginDisconnect");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param connectedShape an input-parameter of type Shape
   * @param connectionSite an input-parameter of type int
   */
  public void endConnect(Shape connectedShape, int connectionSite) {
    Dispatch.call(this, "EndConnect", connectedShape, new Variant(connectionSite));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void endDisconnect() {
    Dispatch.call(this, "EndDisconnect");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBeginConnected() {
    return Dispatch.get(this, "BeginConnected").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Shape
   */
  public Shape getBeginConnectedShape() {
    return new Shape(Dispatch.get(this, "BeginConnectedShape").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBeginConnectionSite() {
    return Dispatch.get(this, "BeginConnectionSite").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getEndConnected() {
    return Dispatch.get(this, "EndConnected").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Shape
   */
  public Shape getEndConnectedShape() {
    return new Shape(Dispatch.get(this, "EndConnectedShape").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getEndConnectionSite() {
    return Dispatch.get(this, "EndConnectionSite").changeType(Variant.VariantInt).getInt();
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
