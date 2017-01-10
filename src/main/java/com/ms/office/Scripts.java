/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Scripts extends Dispatch {

  public static final String componentName = "Office.Scripts";

  public Scripts() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public Scripts(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public Scripts(String compName) {
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
   * @return the result is of type Variant
   */
  public Variant get_NewEnum() {
    return Dispatch.get(this, "_NewEnum");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param index an input-parameter of type Variant
   * @return the result is of type Script
   */
  public Script item(Variant index) {
    return new Script(Dispatch.call(this, "Item", index).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param anchor an input-parameter of type Object
   * @param location an input-parameter of type int
   * @param language an input-parameter of type int
   * @param id an input-parameter of type String
   * @param extended an input-parameter of type String
   * @param scriptText an input-parameter of type String
   * @return the result is of type Script
   */
  public Script add(Object anchor, int location, int language, String id, String extended, String scriptText) {
    return new Script(Dispatch.call(this, "Add", anchor, new Variant(location), new Variant(language), id, extended,
                                    scriptText).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param anchor an input-parameter of type Object
   * @param location an input-parameter of type int
   * @param language an input-parameter of type int
   * @param id an input-parameter of type String
   * @param extended an input-parameter of type String
   * @return the result is of type Script
   */
  public Script add(Object anchor, int location, int language, String id, String extended) {
    return new Script(Dispatch.call(this, "Add", anchor, new Variant(location), new Variant(language), id, extended).
                      toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param anchor an input-parameter of type Object
   * @param location an input-parameter of type int
   * @param language an input-parameter of type int
   * @param id an input-parameter of type String
   * @return the result is of type Script
   */
  public Script add(Object anchor, int location, int language, String id) {
    return new Script(Dispatch.call(this, "Add", anchor, new Variant(location), new Variant(language), id).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param anchor an input-parameter of type Object
   * @param location an input-parameter of type int
   * @param language an input-parameter of type int
   * @return the result is of type Script
   */
  public Script add(Object anchor, int location, int language) {
    return new Script(Dispatch.call(this, "Add", anchor, new Variant(location), new Variant(language)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param anchor an input-parameter of type Object
   * @param location an input-parameter of type int
   * @return the result is of type Script
   */
  public Script add(Object anchor, int location) {
    return new Script(Dispatch.call(this, "Add", anchor, new Variant(location)).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param anchor an input-parameter of type Object
   * @return the result is of type Script
   */
  public Script add(Object anchor) {
    return new Script(Dispatch.call(this, "Add", anchor).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Script
   */
  public Script add() {
    return new Script(Dispatch.call(this, "Add").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void delete() {
    Dispatch.call(this, "Delete");
  }

}
