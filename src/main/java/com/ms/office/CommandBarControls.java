/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class CommandBarControls extends Dispatch {

  public static final String componentName = "Office.CommandBarControls";

  public CommandBarControls() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public CommandBarControls(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public CommandBarControls(String compName) {
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
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @param parameter an input-parameter of type Variant
   * @param before an input-parameter of type Variant
   * @param temporary an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl add(Variant type, Variant id, Variant parameter, Variant before, Variant temporary) {
    return new CommandBarControl(Dispatch.call(this, "Add", type, id, parameter, before, temporary).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @param parameter an input-parameter of type Variant
   * @param before an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl add(Variant type, Variant id, Variant parameter, Variant before) {
    return new CommandBarControl(Dispatch.call(this, "Add", type, id, parameter, before).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @param parameter an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl add(Variant type, Variant id, Variant parameter) {
    return new CommandBarControl(Dispatch.call(this, "Add", type, id, parameter).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl add(Variant type, Variant id) {
    return new CommandBarControl(Dispatch.call(this, "Add", type, id).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl add(Variant type) {
    return new CommandBarControl(Dispatch.call(this, "Add", type).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl add() {
    return new CommandBarControl(Dispatch.call(this, "Add").toDispatch());
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
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl getItem(Variant index) {
    return new CommandBarControl(Dispatch.call(this, "Item", index).toDispatch());
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
   * @return the result is of type CommandBar
   */
  public CommandBar getParent() {
    return new CommandBar(Dispatch.get(this, "Parent").toDispatch());
  }

}
