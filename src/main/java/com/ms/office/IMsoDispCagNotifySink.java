/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class IMsoDispCagNotifySink extends Dispatch {

  public static final String componentName = "Office.IMsoDispCagNotifySink";

  public IMsoDispCagNotifySink() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public IMsoDispCagNotifySink(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public IMsoDispCagNotifySink(String compName) {
    super(compName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param pClipMoniker an input-parameter of type Variant
   * @param pItemMoniker an input-parameter of type Variant
   */
  public void insertClip(Variant pClipMoniker, Variant pItemMoniker) {
    Dispatch.call(this, "InsertClip", pClipMoniker, pItemMoniker);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void windowIsClosing() {
    Dispatch.call(this, "WindowIsClosing");
  }

}
