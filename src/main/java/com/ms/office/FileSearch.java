/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class FileSearch extends Dispatch {

  public static final String componentName = "Office.FileSearch";

  public FileSearch() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public FileSearch(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public FileSearch(String compName) {
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
   * @return the result is of type boolean
   */
  public boolean getSearchSubFolders() {
    return Dispatch.get(this, "SearchSubFolders").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param searchSubFolders an input-parameter of type boolean
   */
  public void setSearchSubFolders(boolean searchSubFolders) {
    Dispatch.put(this, "SearchSubFolders", new Variant(searchSubFolders));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getMatchTextExactly() {
    return Dispatch.get(this, "MatchTextExactly").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param matchTextExactly an input-parameter of type boolean
   */
  public void setMatchTextExactly(boolean matchTextExactly) {
    Dispatch.put(this, "MatchTextExactly", new Variant(matchTextExactly));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getMatchAllWordForms() {
    return Dispatch.get(this, "MatchAllWordForms").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param matchAllWordForms an input-parameter of type boolean
   */
  public void setMatchAllWordForms(boolean matchAllWordForms) {
    Dispatch.put(this, "MatchAllWordForms", new Variant(matchAllWordForms));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getFileName() {
    return Dispatch.get(this, "FileName").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileName an input-parameter of type String
   */
  public void setFileName(String fileName) {
    Dispatch.put(this, "FileName", fileName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getFileType() {
    return Dispatch.get(this, "FileType").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fileType an input-parameter of type int
   */
  public void setFileType(int fileType) {
    Dispatch.put(this, "FileType", new Variant(fileType));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getLastModified() {
    return Dispatch.get(this, "LastModified").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param lastModified an input-parameter of type int
   */
  public void setLastModified(int lastModified) {
    Dispatch.put(this, "LastModified", new Variant(lastModified));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getTextOrProperty() {
    return Dispatch.get(this, "TextOrProperty").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param textOrProperty an input-parameter of type String
   */
  public void setTextOrProperty(String textOrProperty) {
    Dispatch.put(this, "TextOrProperty", textOrProperty);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getLookIn() {
    return Dispatch.get(this, "LookIn").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param lookIn an input-parameter of type String
   */
  public void setLookIn(String lookIn) {
    Dispatch.put(this, "LookIn", lookIn);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param sortBy an input-parameter of type int
   * @param sortOrder an input-parameter of type int
   * @param alwaysAccurate an input-parameter of type boolean
   * @return the result is of type int
   */
  public int execute(int sortBy, int sortOrder, boolean alwaysAccurate) {
    return Dispatch.call(this, "Execute", new Variant(sortBy), new Variant(sortOrder), new Variant(alwaysAccurate)).
            changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param sortBy an input-parameter of type int
   * @param sortOrder an input-parameter of type int
   * @return the result is of type int
   */
  public int execute(int sortBy, int sortOrder) {
    return Dispatch.call(this, "Execute", new Variant(sortBy), new Variant(sortOrder)).changeType(Variant.VariantInt).
            getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param sortBy an input-parameter of type int
   * @return the result is of type int
   */
  public int execute(int sortBy) {
    return Dispatch.call(this, "Execute", new Variant(sortBy)).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int execute() {
    return Dispatch.call(this, "Execute").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void newSearch() {
    Dispatch.call(this, "NewSearch");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type FoundFiles
   */
  public FoundFiles getFoundFiles() {
    return new FoundFiles(Dispatch.get(this, "FoundFiles").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type PropertyTests
   */
  public PropertyTests getPropertyTests() {
    return new PropertyTests(Dispatch.get(this, "PropertyTests").toDispatch());
  }

}
