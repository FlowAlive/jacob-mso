/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class LanguageSettings extends Dispatch {

  public static final String componentName = "Office.LanguageSettings";

  public LanguageSettings() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public LanguageSettings(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public LanguageSettings(String compName) {
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
   * @param id an input-parameter of type int
   * @return the result is of type int
   */
  public int getLanguageID(int id) {
    return Dispatch.call(this, "LanguageID", new Variant(id)).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param lid an input-parameter of type int
   * @return the result is of type boolean
   */
  public boolean getLanguagePreferredForEditing(int lid) {
    return Dispatch.call(this, "LanguagePreferredForEditing", new Variant(lid)).changeType(Variant.VariantBoolean).
            getBoolean();
  }

}
