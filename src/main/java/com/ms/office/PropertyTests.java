/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class PropertyTests extends Dispatch {

  public static final String componentName = "Office.PropertyTests";

  public PropertyTests() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public PropertyTests(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public PropertyTests(String compName) {
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
   * @param index an input-parameter of type int
   * @return the result is of type PropertyTest
   */
  public PropertyTest getItem(int index) {
    return new PropertyTest(Dispatch.call(this, "Item", new Variant(index)).toDispatch());
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
   * @param name an input-parameter of type String
   * @param condition an input-parameter of type int
   * @param value an input-parameter of type Variant
   * @param secondValue an input-parameter of type Variant
   * @param connector an input-parameter of type int
   */
  public void add(String name, int condition, Variant value, Variant secondValue, int connector) {
    Dispatch.call(this, "Add", name, new Variant(condition), value, secondValue, new Variant(connector));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param name an input-parameter of type String
   * @param condition an input-parameter of type int
   * @param value an input-parameter of type Variant
   * @param secondValue an input-parameter of type Variant
   */
  public void add(String name, int condition, Variant value, Variant secondValue) {
    Dispatch.call(this, "Add", name, new Variant(condition), value, secondValue);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param name an input-parameter of type String
   * @param condition an input-parameter of type int
   * @param value an input-parameter of type Variant
   */
  public void add(String name, int condition, Variant value) {
    Dispatch.call(this, "Add", name, new Variant(condition), value);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param name an input-parameter of type String
   * @param condition an input-parameter of type int
   */
  public void add(String name, int condition) {
    Dispatch.call(this, "Add", name, new Variant(condition));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param index an input-parameter of type int
   */
  public void remove(int index) {
    Dispatch.call(this, "Remove", new Variant(index));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant get_NewEnum() {
    return Dispatch.get(this, "_NewEnum");
  }

}
