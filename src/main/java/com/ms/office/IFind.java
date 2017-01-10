/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class IFind extends Dispatch {

  public static final String componentName = "Office.IFind";

  public IFind() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public IFind(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public IFind(String compName) {
    super(compName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getSearchPath() {
    return Dispatch.get(this, "SearchPath").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getName() {
    return Dispatch.get(this, "Name").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getSubDir() {
    return Dispatch.get(this, "SubDir").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getTitle() {
    return Dispatch.get(this, "Title").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getAuthor() {
    return Dispatch.get(this, "Author").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getKeywords() {
    return Dispatch.get(this, "Keywords").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getSubject() {
    return Dispatch.get(this, "Subject").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getOptions() {
    return Dispatch.get(this, "Options").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getMatchCase() {
    return Dispatch.get(this, "MatchCase").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getText() {
    return Dispatch.get(this, "Text").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getPatternMatch() {
    return Dispatch.get(this, "PatternMatch").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant getDateSavedFrom() {
    return Dispatch.get(this, "DateSavedFrom");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant getDateSavedTo() {
    return Dispatch.get(this, "DateSavedTo");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getSavedBy() {
    return Dispatch.get(this, "SavedBy").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant getDateCreatedFrom() {
    return Dispatch.get(this, "DateCreatedFrom");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant getDateCreatedTo() {
    return Dispatch.get(this, "DateCreatedTo");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getView() {
    return Dispatch.get(this, "View").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getSortBy() {
    return Dispatch.get(this, "SortBy").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getListBy() {
    return Dispatch.get(this, "ListBy").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getSelectedFile() {
    return Dispatch.get(this, "SelectedFile").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type IFoundFiles
   */
  public IFoundFiles getResults() {
    return new IFoundFiles(Dispatch.get(this, "Results").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int show() {
    return Dispatch.call(this, "Show").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param searchPath an input-parameter of type String
   */
  public void setSearchPath(String searchPath) {
    Dispatch.put(this, "SearchPath", searchPath);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param name an input-parameter of type String
   */
  public void setName(String name) {
    Dispatch.put(this, "Name", name);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param subDir an input-parameter of type boolean
   */
  public void setSubDir(boolean subDir) {
    Dispatch.put(this, "SubDir", new Variant(subDir));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param title an input-parameter of type String
   */
  public void setTitle(String title) {
    Dispatch.put(this, "Title", title);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param author an input-parameter of type String
   */
  public void setAuthor(String author) {
    Dispatch.put(this, "Author", author);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param keywords an input-parameter of type String
   */
  public void setKeywords(String keywords) {
    Dispatch.put(this, "Keywords", keywords);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param subject an input-parameter of type String
   */
  public void setSubject(String subject) {
    Dispatch.put(this, "Subject", subject);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param options an input-parameter of type int
   */
  public void setOptions(int options) {
    Dispatch.put(this, "Options", new Variant(options));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param matchCase an input-parameter of type boolean
   */
  public void setMatchCase(boolean matchCase) {
    Dispatch.put(this, "MatchCase", new Variant(matchCase));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param text an input-parameter of type String
   */
  public void setText(String text) {
    Dispatch.put(this, "Text", text);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param patternMatch an input-parameter of type boolean
   */
  public void setPatternMatch(boolean patternMatch) {
    Dispatch.put(this, "PatternMatch", new Variant(patternMatch));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param dateSavedFrom an input-parameter of type Variant
   */
  public void setDateSavedFrom(Variant dateSavedFrom) {
    Dispatch.put(this, "DateSavedFrom", dateSavedFrom);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param dateSavedTo an input-parameter of type Variant
   */
  public void setDateSavedTo(Variant dateSavedTo) {
    Dispatch.put(this, "DateSavedTo", dateSavedTo);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param savedBy an input-parameter of type String
   */
  public void setSavedBy(String savedBy) {
    Dispatch.put(this, "SavedBy", savedBy);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param dateCreatedFrom an input-parameter of type Variant
   */
  public void setDateCreatedFrom(Variant dateCreatedFrom) {
    Dispatch.put(this, "DateCreatedFrom", dateCreatedFrom);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param dateCreatedTo an input-parameter of type Variant
   */
  public void setDateCreatedTo(Variant dateCreatedTo) {
    Dispatch.put(this, "DateCreatedTo", dateCreatedTo);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param view an input-parameter of type int
   */
  public void setView(int view) {
    Dispatch.put(this, "View", new Variant(view));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param sortBy an input-parameter of type int
   */
  public void setSortBy(int sortBy) {
    Dispatch.put(this, "SortBy", new Variant(sortBy));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param listBy an input-parameter of type int
   */
  public void setListBy(int listBy) {
    Dispatch.put(this, "ListBy", new Variant(listBy));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param selectedFile an input-parameter of type int
   */
  public void setSelectedFile(int selectedFile) {
    Dispatch.put(this, "SelectedFile", new Variant(selectedFile));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void execute() {
    Dispatch.call(this, "Execute");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param bstrQueryName an input-parameter of type String
   */
  public void load(String bstrQueryName) {
    Dispatch.call(this, "Load", bstrQueryName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param bstrQueryName an input-parameter of type String
   */
  public void save(String bstrQueryName) {
    Dispatch.call(this, "Save", bstrQueryName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param bstrQueryName an input-parameter of type String
   */
  public void delete(String bstrQueryName) {
    Dispatch.call(this, "Delete", bstrQueryName);
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

}
