/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ICommandBarButtonEvents extends Dispatch {

  public static final String componentName = "Office.ICommandBarButtonEvents";

  public ICommandBarButtonEvents() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public ICommandBarButtonEvents(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public ICommandBarButtonEvents(String compName) {
    super(compName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param ctrl an input-parameter of type CommandBarButton
   * @param cancelDefault an input-parameter of type boolean
   */
  public void click(CommandBarButton ctrl, boolean cancelDefault) {
    Dispatch.call(this, "Click", ctrl, new Variant(cancelDefault));
  }

  /**
   * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
   * @param ctrl an input-parameter of type CommandBarButton
   * @param cancelDefault is an one-element array which sends the input-parameter
   *                      to the ActiveX-Component and receives the output-parameter
   */
  public void click(CommandBarButton ctrl, boolean[] cancelDefault) {
    Variant vnt_cancelDefault = new Variant();
    if (cancelDefault == null || cancelDefault.length == 0) {
      vnt_cancelDefault.putNoParam();
    }
    else {
      vnt_cancelDefault.putBooleanRef(cancelDefault[0]);
    }

    Dispatch.call(this, "Click", ctrl, vnt_cancelDefault);

    if (cancelDefault != null && cancelDefault.length > 0) {
      cancelDefault[0] = vnt_cancelDefault.toBoolean();
    }
  }

}
