/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Documents extends Dispatch {

  public static final String componentName = "Word.Documents";

  public Documents() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public Documents(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public Documents(String compName) {
    super(compName);
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
   * @return the result is of type int
   */
  public int getCount() {
    return Dispatch.get(this, "Count").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Application
   */
  public AppEvents getApplication() {
    return new AppEvents(Dispatch.get(this, "Application").toDispatch());
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
   * @param index an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document item(Variant index) {
    return new Document(Dispatch.call(this, "Item", index).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param saveChanges an input-parameter of type Variant
   * @param originalFormat an input-parameter of type Variant
   * @param routeDocument an input-parameter of type Variant
   */
  public void close(Variant saveChanges, Variant originalFormat, Variant routeDocument) {
    Dispatch.call(this, "Close", saveChanges, originalFormat, routeDocument);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param saveChanges an input-parameter of type Variant
   * @param originalFormat an input-parameter of type Variant
   */
  public void close(Variant saveChanges, Variant originalFormat) {
    Dispatch.call(this, "Close", saveChanges, originalFormat);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param saveChanges an input-parameter of type Variant
   */
  public void close(Variant saveChanges) {
    Dispatch.call(this, "Close", saveChanges);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void close() {
    Dispatch.call(this, "Close");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param template an input-parameter of type Variant
   * @param newTemplate an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document addOld(Variant template, Variant newTemplate) {
    return new Document(Dispatch.call(this, "AddOld", template, newTemplate).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param template an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document addOld(Variant template) {
    return new Document(Dispatch.call(this, "AddOld", template).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Document
   */
  public Document addOld() {
    return new Document(Dispatch.call(this, "AddOld").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @param writePasswordDocument an input-parameter of type Variant
   * @param writePasswordTemplate an input-parameter of type Variant
   * @param format an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                          Variant passwordDocument, Variant passwordTemplate, Variant revert,
                          Variant writePasswordDocument, Variant writePasswordTemplate, Variant format) {
    return new Document(Dispatch.callN(this, "OpenOld", new Object[] {fileName, confirmConversions, readOnly,
                                       addToRecentFiles, passwordDocument, passwordTemplate, revert,
                                       writePasswordDocument, writePasswordTemplate, format}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @param writePasswordDocument an input-parameter of type Variant
   * @param writePasswordTemplate an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                          Variant passwordDocument, Variant passwordTemplate, Variant revert,
                          Variant writePasswordDocument, Variant writePasswordTemplate) {
    return new Document(Dispatch.callN(this, "OpenOld", new Object[] {fileName, confirmConversions, readOnly,
                                       addToRecentFiles, passwordDocument, passwordTemplate, revert,
                                       writePasswordDocument, writePasswordTemplate}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @param writePasswordDocument an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                          Variant passwordDocument, Variant passwordTemplate, Variant revert,
                          Variant writePasswordDocument) {
    return new Document(Dispatch.call(this, "OpenOld", fileName, confirmConversions, readOnly, addToRecentFiles,
                                      passwordDocument, passwordTemplate, revert, writePasswordDocument).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                          Variant passwordDocument, Variant passwordTemplate, Variant revert) {
    return new Document(Dispatch.call(this, "OpenOld", fileName, confirmConversions, readOnly, addToRecentFiles,
                                      passwordDocument, passwordTemplate, revert).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                          Variant passwordDocument, Variant passwordTemplate) {
    return new Document(Dispatch.call(this, "OpenOld", fileName, confirmConversions, readOnly, addToRecentFiles,
                                      passwordDocument, passwordTemplate).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                          Variant passwordDocument) {
    return new Document(Dispatch.call(this, "OpenOld", fileName, confirmConversions, readOnly, addToRecentFiles,
                                      passwordDocument).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles) {
    return new Document(Dispatch.call(this, "OpenOld", fileName, confirmConversions, readOnly, addToRecentFiles).
                        toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName, Variant confirmConversions, Variant readOnly) {
    return new Document(Dispatch.call(this, "OpenOld", fileName, confirmConversions, readOnly).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName, Variant confirmConversions) {
    return new Document(Dispatch.call(this, "OpenOld", fileName, confirmConversions).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document openOld(Variant fileName) {
    return new Document(Dispatch.call(this, "OpenOld", fileName).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param noPrompt an input-parameter of type Variant
   * @param originalFormat an input-parameter of type Variant
   */
  public void save(Variant noPrompt, Variant originalFormat) {
    Dispatch.call(this, "Save", noPrompt, originalFormat);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param noPrompt an input-parameter of type Variant
   */
  public void save(Variant noPrompt) {
    Dispatch.call(this, "Save", noPrompt);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void save() {
    Dispatch.call(this, "Save");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param template an input-parameter of type Variant
   * @param newTemplate an input-parameter of type Variant
   * @param documentType an input-parameter of type Variant
   * @param visible an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document add(Variant template, Variant newTemplate, Variant documentType, Variant visible) {
    return new Document(Dispatch.call(this, "Add", template, newTemplate, documentType, visible).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param template an input-parameter of type Variant
   * @param newTemplate an input-parameter of type Variant
   * @param documentType an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document add(Variant template, Variant newTemplate, Variant documentType) {
    return new Document(Dispatch.call(this, "Add", template, newTemplate, documentType).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param template an input-parameter of type Variant
   * @param newTemplate an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document add(Variant template, Variant newTemplate) {
    return new Document(Dispatch.call(this, "Add", template, newTemplate).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param template an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document add(Variant template) {
    return new Document(Dispatch.call(this, "Add", template).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Document
   */
  public Document add() {
    return new Document(Dispatch.call(this, "Add").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @param writePasswordDocument an input-parameter of type Variant
   * @param writePasswordTemplate an input-parameter of type Variant
   * @param format an input-parameter of type Variant
   * @param encoding an input-parameter of type Variant
   * @param visible an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                       Variant passwordDocument, Variant passwordTemplate, Variant revert,
                       Variant writePasswordDocument, Variant writePasswordTemplate, Variant format, Variant encoding,
                       Variant visible) {
    return new Document(Dispatch.callN(this, "Open", new Object[] {fileName, confirmConversions, readOnly,
                                       addToRecentFiles, passwordDocument, passwordTemplate, revert,
                                       writePasswordDocument, writePasswordTemplate, format, encoding, visible}).
                        toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @param writePasswordDocument an input-parameter of type Variant
   * @param writePasswordTemplate an input-parameter of type Variant
   * @param format an input-parameter of type Variant
   * @param encoding an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                       Variant passwordDocument, Variant passwordTemplate, Variant revert,
                       Variant writePasswordDocument, Variant writePasswordTemplate, Variant format, Variant encoding) {
    return new Document(Dispatch.callN(this, "Open", new Object[] {fileName, confirmConversions, readOnly,
                                       addToRecentFiles, passwordDocument, passwordTemplate, revert,
                                       writePasswordDocument, writePasswordTemplate, format, encoding}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @param writePasswordDocument an input-parameter of type Variant
   * @param writePasswordTemplate an input-parameter of type Variant
   * @param format an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                       Variant passwordDocument, Variant passwordTemplate, Variant revert,
                       Variant writePasswordDocument, Variant writePasswordTemplate, Variant format) {
    return new Document(Dispatch.callN(this, "Open", new Object[] {fileName, confirmConversions, readOnly,
                                       addToRecentFiles, passwordDocument, passwordTemplate, revert,
                                       writePasswordDocument, writePasswordTemplate, format}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @param writePasswordDocument an input-parameter of type Variant
   * @param writePasswordTemplate an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                       Variant passwordDocument, Variant passwordTemplate, Variant revert,
                       Variant writePasswordDocument, Variant writePasswordTemplate) {
    return new Document(Dispatch.callN(this, "Open", new Object[] {fileName, confirmConversions, readOnly,
                                       addToRecentFiles, passwordDocument, passwordTemplate, revert,
                                       writePasswordDocument, writePasswordTemplate}).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @param writePasswordDocument an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                       Variant passwordDocument, Variant passwordTemplate, Variant revert,
                       Variant writePasswordDocument) {
    return new Document(Dispatch.call(this, "Open", fileName, confirmConversions, readOnly, addToRecentFiles,
                                      passwordDocument, passwordTemplate, revert, writePasswordDocument).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @param revert an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                       Variant passwordDocument, Variant passwordTemplate, Variant revert) {
    return new Document(Dispatch.call(this, "Open", fileName, confirmConversions, readOnly, addToRecentFiles,
                                      passwordDocument, passwordTemplate, revert).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @param passwordTemplate an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                       Variant passwordDocument, Variant passwordTemplate) {
    return new Document(Dispatch.call(this, "Open", fileName, confirmConversions, readOnly, addToRecentFiles,
                                      passwordDocument, passwordTemplate).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @param passwordDocument an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles,
                       Variant passwordDocument) {
    return new Document(Dispatch.call(this, "Open", fileName, confirmConversions, readOnly, addToRecentFiles,
                                      passwordDocument).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @param addToRecentFiles an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly, Variant addToRecentFiles) {
    return new Document(Dispatch.call(this, "Open", fileName, confirmConversions, readOnly, addToRecentFiles).
                        toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @param readOnly an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions, Variant readOnly) {
    return new Document(Dispatch.call(this, "Open", fileName, confirmConversions, readOnly).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @param confirmConversions an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName, Variant confirmConversions) {
    return new Document(Dispatch.call(this, "Open", fileName, confirmConversions).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type Variant
   * @return the result is of type Document
   */
  public Document open(Variant fileName) {
    return new Document(Dispatch.call(this, "Open", fileName).toDispatch());
  }

}
