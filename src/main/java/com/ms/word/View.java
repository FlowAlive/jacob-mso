/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class View extends Dispatch {

  public static final String componentName = "Word.View";

  public View() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public View(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public View(String compName) {
    super(compName);
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
   * @return the result is of type int
   */
  public int getType() {
    return Dispatch.get(this, "Type").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type int
   */
  public void setType(int type) {
    Dispatch.put(this, "Type", new Variant(type));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getFullScreen() {
    return Dispatch.get(this, "FullScreen").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fullScreen an input-parameter of type boolean
   */
  public void setFullScreen(boolean fullScreen) {
    Dispatch.put(this, "FullScreen", new Variant(fullScreen));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getDraft() {
    return Dispatch.get(this, "Draft").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param draft an input-parameter of type boolean
   */
  public void setDraft(boolean draft) {
    Dispatch.put(this, "Draft", new Variant(draft));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowAll() {
    return Dispatch.get(this, "ShowAll").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showAll an input-parameter of type boolean
   */
  public void setShowAll(boolean showAll) {
    Dispatch.put(this, "ShowAll", new Variant(showAll));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowFieldCodes() {
    return Dispatch.get(this, "ShowFieldCodes").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showFieldCodes an input-parameter of type boolean
   */
  public void setShowFieldCodes(boolean showFieldCodes) {
    Dispatch.put(this, "ShowFieldCodes", new Variant(showFieldCodes));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getMailMergeDataView() {
    return Dispatch.get(this, "MailMergeDataView").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param mailMergeDataView an input-parameter of type boolean
   */
  public void setMailMergeDataView(boolean mailMergeDataView) {
    Dispatch.put(this, "MailMergeDataView", new Variant(mailMergeDataView));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getMagnifier() {
    return Dispatch.get(this, "Magnifier").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param magnifier an input-parameter of type boolean
   */
  public void setMagnifier(boolean magnifier) {
    Dispatch.put(this, "Magnifier", new Variant(magnifier));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowFirstLineOnly() {
    return Dispatch.get(this, "ShowFirstLineOnly").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showFirstLineOnly an input-parameter of type boolean
   */
  public void setShowFirstLineOnly(boolean showFirstLineOnly) {
    Dispatch.put(this, "ShowFirstLineOnly", new Variant(showFirstLineOnly));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowFormat() {
    return Dispatch.get(this, "ShowFormat").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showFormat an input-parameter of type boolean
   */
  public void setShowFormat(boolean showFormat) {
    Dispatch.put(this, "ShowFormat", new Variant(showFormat));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Zoom
   */
  public Zoom getZoom() {
    return new Zoom(Dispatch.get(this, "Zoom").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowObjectAnchors() {
    return Dispatch.get(this, "ShowObjectAnchors").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showObjectAnchors an input-parameter of type boolean
   */
  public void setShowObjectAnchors(boolean showObjectAnchors) {
    Dispatch.put(this, "ShowObjectAnchors", new Variant(showObjectAnchors));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowTextBoundaries() {
    return Dispatch.get(this, "ShowTextBoundaries").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showTextBoundaries an input-parameter of type boolean
   */
  public void setShowTextBoundaries(boolean showTextBoundaries) {
    Dispatch.put(this, "ShowTextBoundaries", new Variant(showTextBoundaries));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowHighlight() {
    return Dispatch.get(this, "ShowHighlight").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showHighlight an input-parameter of type boolean
   */
  public void setShowHighlight(boolean showHighlight) {
    Dispatch.put(this, "ShowHighlight", new Variant(showHighlight));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowDrawings() {
    return Dispatch.get(this, "ShowDrawings").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showDrawings an input-parameter of type boolean
   */
  public void setShowDrawings(boolean showDrawings) {
    Dispatch.put(this, "ShowDrawings", new Variant(showDrawings));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowTabs() {
    return Dispatch.get(this, "ShowTabs").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showTabs an input-parameter of type boolean
   */
  public void setShowTabs(boolean showTabs) {
    Dispatch.put(this, "ShowTabs", new Variant(showTabs));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowSpaces() {
    return Dispatch.get(this, "ShowSpaces").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showSpaces an input-parameter of type boolean
   */
  public void setShowSpaces(boolean showSpaces) {
    Dispatch.put(this, "ShowSpaces", new Variant(showSpaces));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowParagraphs() {
    return Dispatch.get(this, "ShowParagraphs").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showParagraphs an input-parameter of type boolean
   */
  public void setShowParagraphs(boolean showParagraphs) {
    Dispatch.put(this, "ShowParagraphs", new Variant(showParagraphs));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowHyphens() {
    return Dispatch.get(this, "ShowHyphens").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showHyphens an input-parameter of type boolean
   */
  public void setShowHyphens(boolean showHyphens) {
    Dispatch.put(this, "ShowHyphens", new Variant(showHyphens));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowHiddenText() {
    return Dispatch.get(this, "ShowHiddenText").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showHiddenText an input-parameter of type boolean
   */
  public void setShowHiddenText(boolean showHiddenText) {
    Dispatch.put(this, "ShowHiddenText", new Variant(showHiddenText));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getWrapToWindow() {
    return Dispatch.get(this, "WrapToWindow").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param wrapToWindow an input-parameter of type boolean
   */
  public void setWrapToWindow(boolean wrapToWindow) {
    Dispatch.put(this, "WrapToWindow", new Variant(wrapToWindow));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowPicturePlaceHolders() {
    return Dispatch.get(this, "ShowPicturePlaceHolders").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showPicturePlaceHolders an input-parameter of type boolean
   */
  public void setShowPicturePlaceHolders(boolean showPicturePlaceHolders) {
    Dispatch.put(this, "ShowPicturePlaceHolders", new Variant(showPicturePlaceHolders));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowBookmarks() {
    return Dispatch.get(this, "ShowBookmarks").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showBookmarks an input-parameter of type boolean
   */
  public void setShowBookmarks(boolean showBookmarks) {
    Dispatch.put(this, "ShowBookmarks", new Variant(showBookmarks));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getFieldShading() {
    return Dispatch.get(this, "FieldShading").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param fieldShading an input-parameter of type int
   */
  public void setFieldShading(int fieldShading) {
    Dispatch.put(this, "FieldShading", new Variant(fieldShading));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowAnimation() {
    return Dispatch.get(this, "ShowAnimation").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showAnimation an input-parameter of type boolean
   */
  public void setShowAnimation(boolean showAnimation) {
    Dispatch.put(this, "ShowAnimation", new Variant(showAnimation));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getTableGridlines() {
    return Dispatch.get(this, "TableGridlines").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param tableGridlines an input-parameter of type boolean
   */
  public void setTableGridlines(boolean tableGridlines) {
    Dispatch.put(this, "TableGridlines", new Variant(tableGridlines));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getEnlargeFontsLessThan() {
    return Dispatch.get(this, "EnlargeFontsLessThan").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param enlargeFontsLessThan an input-parameter of type int
   */
  public void setEnlargeFontsLessThan(int enlargeFontsLessThan) {
    Dispatch.put(this, "EnlargeFontsLessThan", new Variant(enlargeFontsLessThan));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowMainTextLayer() {
    return Dispatch.get(this, "ShowMainTextLayer").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showMainTextLayer an input-parameter of type boolean
   */
  public void setShowMainTextLayer(boolean showMainTextLayer) {
    Dispatch.put(this, "ShowMainTextLayer", new Variant(showMainTextLayer));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getSeekView() {
    return Dispatch.get(this, "SeekView").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param seekView an input-parameter of type int
   */
  public void setSeekView(int seekView) {
    Dispatch.put(this, "SeekView", new Variant(seekView));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getSplitSpecial() {
    return Dispatch.get(this, "SplitSpecial").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param splitSpecial an input-parameter of type int
   */
  public void setSplitSpecial(int splitSpecial) {
    Dispatch.put(this, "SplitSpecial", new Variant(splitSpecial));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBrowseToWindow() {
    return Dispatch.get(this, "BrowseToWindow").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param browseToWindow an input-parameter of type int
   */
  public void setBrowseToWindow(int browseToWindow) {
    Dispatch.put(this, "BrowseToWindow", new Variant(browseToWindow));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getShowOptionalBreaks() {
    return Dispatch.get(this, "ShowOptionalBreaks").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param showOptionalBreaks an input-parameter of type boolean
   */
  public void setShowOptionalBreaks(boolean showOptionalBreaks) {
    Dispatch.put(this, "ShowOptionalBreaks", new Variant(showOptionalBreaks));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param range an input-parameter of type Variant
   */
  public void collapseOutline(Variant range) {
    Dispatch.call(this, "CollapseOutline", range);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void collapseOutline() {
    Dispatch.call(this, "CollapseOutline");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param range an input-parameter of type Variant
   */
  public void expandOutline(Variant range) {
    Dispatch.call(this, "ExpandOutline", range);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void expandOutline() {
    Dispatch.call(this, "ExpandOutline");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void showAllHeadings() {
    Dispatch.call(this, "ShowAllHeadings");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param level an input-parameter of type int
   */
  public void showHeading(int level) {
    Dispatch.call(this, "ShowHeading", new Variant(level));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void previousHeaderFooter() {
    Dispatch.call(this, "PreviousHeaderFooter");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void nextHeaderFooter() {
    Dispatch.call(this, "NextHeaderFooter");
  }

}
