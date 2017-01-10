/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;

public class _CommandBarComboBoxEvents extends Dispatch {

  public static final String componentName = "Office._CommandBarComboBoxEvents";

  public _CommandBarComboBoxEvents() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public _CommandBarComboBoxEvents(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public _CommandBarComboBoxEvents(String compName) {
    super(compName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param ctrl an input-parameter of type CommandBarComboBox
   */
  public void change(CommandBarComboBox ctrl) {
    Dispatch.call(this, "Change", ctrl);
  }

}
