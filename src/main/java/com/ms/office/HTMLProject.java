/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class HTMLProject extends Dispatch {

  public static final String componentName = "Office.HTMLProject";

  public HTMLProject() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public HTMLProject(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public HTMLProject(String compName) {
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
   * @return the result is of type int
   */
  public int getState() {
    return Dispatch.get(this, "State").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param refresh an input-parameter of type boolean
   */
  public void refreshProject(boolean refresh) {
    Dispatch.call(this, "RefreshProject", new Variant(refresh));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void refreshProject() {
    Dispatch.call(this, "RefreshProject");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param refresh an input-parameter of type boolean
   */
  public void refreshDocument(boolean refresh) {
    Dispatch.call(this, "RefreshDocument", new Variant(refresh));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void refreshDocument() {
    Dispatch.call(this, "RefreshDocument");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type HTMLProjectItems
   */
  public HTMLProjectItems getHTMLProjectItems() {
    return new HTMLProjectItems(Dispatch.get(this, "HTMLProjectItems").toDispatch());
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
   * @param openKind an input-parameter of type int
   */
  public void open(int openKind) {
    Dispatch.call(this, "Open", new Variant(openKind));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void open() {
    Dispatch.call(this, "Open");
  }

}
