/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class AnswerWizard extends Dispatch {

  public static final String componentName = "Office.AnswerWizard";

  public AnswerWizard() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public AnswerWizard(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public AnswerWizard(String compName) {
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
   * @return the result is of type AnswerWizardFiles
   */
  public AnswerWizardFiles getFiles() {
    return new AnswerWizardFiles(Dispatch.get(this, "Files").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void clearFileList() {
    Dispatch.call(this, "ClearFileList");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void resetFileList() {
    Dispatch.call(this, "ResetFileList");
  }

}
