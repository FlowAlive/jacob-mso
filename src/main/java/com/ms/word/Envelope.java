/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Envelope extends Dispatch {

  public static final String componentName = "Word.Envelope";

  public Envelope() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public Envelope(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public Envelope(String compName) {
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
   * @return the result is of type Range
   */
  public Range getAddress() {
    return new Range(Dispatch.get(this, "Address").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Range
   */
  public Range getReturnAddress() {
    return new Range(Dispatch.get(this, "ReturnAddress").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getDefaultPrintBarCode() {
    return Dispatch.get(this, "DefaultPrintBarCode").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param defaultPrintBarCode an input-parameter of type boolean
   */
  public void setDefaultPrintBarCode(boolean defaultPrintBarCode) {
    Dispatch.put(this, "DefaultPrintBarCode", new Variant(defaultPrintBarCode));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getDefaultPrintFIMA() {
    return Dispatch.get(this, "DefaultPrintFIMA").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param defaultPrintFIMA an input-parameter of type boolean
   */
  public void setDefaultPrintFIMA(boolean defaultPrintFIMA) {
    Dispatch.put(this, "DefaultPrintFIMA", new Variant(defaultPrintFIMA));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getDefaultHeight() {
    return Dispatch.get(this, "DefaultHeight").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param defaultHeight an input-parameter of type float
   */
  public void setDefaultHeight(float defaultHeight) {
    Dispatch.put(this, "DefaultHeight", new Variant(defaultHeight));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getDefaultWidth() {
    return Dispatch.get(this, "DefaultWidth").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param defaultWidth an input-parameter of type float
   */
  public void setDefaultWidth(float defaultWidth) {
    Dispatch.put(this, "DefaultWidth", new Variant(defaultWidth));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getDefaultSize() {
    return Dispatch.get(this, "DefaultSize").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param defaultSize an input-parameter of type String
   */
  public void setDefaultSize(String defaultSize) {
    Dispatch.put(this, "DefaultSize", defaultSize);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getDefaultOmitReturnAddress() {
    return Dispatch.get(this, "DefaultOmitReturnAddress").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param defaultOmitReturnAddress an input-parameter of type boolean
   */
  public void setDefaultOmitReturnAddress(boolean defaultOmitReturnAddress) {
    Dispatch.put(this, "DefaultOmitReturnAddress", new Variant(defaultOmitReturnAddress));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getFeedSource() {
    return Dispatch.get(this, "FeedSource").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param feedSource an input-parameter of type int
   */
  public void setFeedSource(int feedSource) {
    Dispatch.put(this, "FeedSource", new Variant(feedSource));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getAddressFromLeft() {
    return Dispatch.get(this, "AddressFromLeft").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param addressFromLeft an input-parameter of type float
   */
  public void setAddressFromLeft(float addressFromLeft) {
    Dispatch.put(this, "AddressFromLeft", new Variant(addressFromLeft));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getAddressFromTop() {
    return Dispatch.get(this, "AddressFromTop").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param addressFromTop an input-parameter of type float
   */
  public void setAddressFromTop(float addressFromTop) {
    Dispatch.put(this, "AddressFromTop", new Variant(addressFromTop));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getReturnAddressFromLeft() {
    return Dispatch.get(this, "ReturnAddressFromLeft").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param returnAddressFromLeft an input-parameter of type float
   */
  public void setReturnAddressFromLeft(float returnAddressFromLeft) {
    Dispatch.put(this, "ReturnAddressFromLeft", new Variant(returnAddressFromLeft));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type float
   */
  public float getReturnAddressFromTop() {
    return Dispatch.get(this, "ReturnAddressFromTop").changeType(Variant.VariantFloat).getFloat();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param returnAddressFromTop an input-parameter of type float
   */
  public void setReturnAddressFromTop(float returnAddressFromTop) {
    Dispatch.put(this, "ReturnAddressFromTop", new Variant(returnAddressFromTop));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Style
   */
  public Style getAddressStyle() {
    return new Style(Dispatch.get(this, "AddressStyle").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Style
   */
  public Style getReturnAddressStyle() {
    return new Style(Dispatch.get(this, "ReturnAddressStyle").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getDefaultOrientation() {
    return Dispatch.get(this, "DefaultOrientation").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param defaultOrientation an input-parameter of type int
   */
  public void setDefaultOrientation(int defaultOrientation) {
    Dispatch.put(this, "DefaultOrientation", new Variant(defaultOrientation));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getDefaultFaceUp() {
    return Dispatch.get(this, "DefaultFaceUp").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param defaultFaceUp an input-parameter of type boolean
   */
  public void setDefaultFaceUp(boolean defaultFaceUp) {
    Dispatch.put(this, "DefaultFaceUp", new Variant(defaultFaceUp));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   * @param returnAddressFromLeft an input-parameter of type Variant
   * @param returnAddressFromTop an input-parameter of type Variant
   * @param defaultFaceUp an input-parameter of type Variant
   * @param defaultOrientation an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                     Variant addressFromTop, Variant returnAddressFromLeft, Variant returnAddressFromTop,
                     Variant defaultFaceUp, Variant defaultOrientation) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   * @param returnAddressFromLeft an input-parameter of type Variant
   * @param returnAddressFromTop an input-parameter of type Variant
   * @param defaultFaceUp an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                     Variant addressFromTop, Variant returnAddressFromLeft, Variant returnAddressFromTop,
                     Variant defaultFaceUp) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   * @param returnAddressFromLeft an input-parameter of type Variant
   * @param returnAddressFromTop an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                     Variant addressFromTop, Variant returnAddressFromLeft, Variant returnAddressFromTop) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop, returnAddressFromLeft, returnAddressFromTop});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   * @param returnAddressFromLeft an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                     Variant addressFromTop, Variant returnAddressFromLeft) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop, returnAddressFromLeft});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                     Variant addressFromTop) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size, Variant height, Variant width, Variant feedSource) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size, Variant height, Variant width) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size, Variant height) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                     Variant size) {
    Dispatch.callN(this, "Insert", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA) {
    Dispatch.call(this, "Insert", extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText,
                  printBarCode, printFIMA);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText, Variant printBarCode) {
    Dispatch.call(this, "Insert", extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText,
                  printBarCode);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress, Variant returnAutoText) {
    Dispatch.call(this, "Insert", extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                     Variant returnAddress) {
    Dispatch.call(this, "Insert", extractAddress, address, autoText, omitReturnAddress, returnAddress);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress) {
    Dispatch.call(this, "Insert", extractAddress, address, autoText, omitReturnAddress);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address, Variant autoText) {
    Dispatch.call(this, "Insert", extractAddress, address, autoText);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   */
  public void insert(Variant extractAddress, Variant address) {
    Dispatch.call(this, "Insert", extractAddress, address);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   */
  public void insert(Variant extractAddress) {
    Dispatch.call(this, "Insert", extractAddress);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void insert() {
    Dispatch.call(this, "Insert");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   * @param returnAddressFromLeft an input-parameter of type Variant
   * @param returnAddressFromTop an input-parameter of type Variant
   * @param defaultFaceUp an input-parameter of type Variant
   * @param defaultOrientation an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                       Variant addressFromTop, Variant returnAddressFromLeft, Variant returnAddressFromTop,
                       Variant defaultFaceUp, Variant defaultOrientation) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp, defaultOrientation});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   * @param returnAddressFromLeft an input-parameter of type Variant
   * @param returnAddressFromTop an input-parameter of type Variant
   * @param defaultFaceUp an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                       Variant addressFromTop, Variant returnAddressFromLeft, Variant returnAddressFromTop,
                       Variant defaultFaceUp) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop, returnAddressFromLeft, returnAddressFromTop, defaultFaceUp});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   * @param returnAddressFromLeft an input-parameter of type Variant
   * @param returnAddressFromTop an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                       Variant addressFromTop, Variant returnAddressFromLeft, Variant returnAddressFromTop) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop, returnAddressFromLeft, returnAddressFromTop});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   * @param returnAddressFromLeft an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                       Variant addressFromTop, Variant returnAddressFromLeft) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop, returnAddressFromLeft});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   * @param addressFromTop an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft,
                       Variant addressFromTop) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft,
                   addressFromTop});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   * @param addressFromLeft an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size, Variant height, Variant width, Variant feedSource, Variant addressFromLeft) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource, addressFromLeft});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   * @param feedSource an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size, Variant height, Variant width, Variant feedSource) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width, feedSource});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   * @param width an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size, Variant height, Variant width) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height, width});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   * @param height an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size, Variant height) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size, height});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   * @param size an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA,
                       Variant size) {
    Dispatch.callN(this, "PrintOut", new Object[] {extractAddress, address, autoText, omitReturnAddress, returnAddress,
                   returnAutoText, printBarCode, printFIMA, size});
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   * @param printFIMA an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode, Variant printFIMA) {
    Dispatch.call(this, "PrintOut", extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText,
                  printBarCode, printFIMA);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   * @param printBarCode an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText, Variant printBarCode) {
    Dispatch.call(this, "PrintOut", extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText,
                  printBarCode);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   * @param returnAutoText an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress, Variant returnAutoText) {
    Dispatch.call(this, "PrintOut", extractAddress, address, autoText, omitReturnAddress, returnAddress, returnAutoText);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   * @param returnAddress an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress,
                       Variant returnAddress) {
    Dispatch.call(this, "PrintOut", extractAddress, address, autoText, omitReturnAddress, returnAddress);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   * @param omitReturnAddress an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText, Variant omitReturnAddress) {
    Dispatch.call(this, "PrintOut", extractAddress, address, autoText, omitReturnAddress);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   * @param autoText an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address, Variant autoText) {
    Dispatch.call(this, "PrintOut", extractAddress, address, autoText);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   * @param address an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress, Variant address) {
    Dispatch.call(this, "PrintOut", extractAddress, address);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param extractAddress an input-parameter of type Variant
   */
  public void printOut(Variant extractAddress) {
    Dispatch.call(this, "PrintOut", extractAddress);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void printOut() {
    Dispatch.call(this, "PrintOut");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void updateDocument() {
    Dispatch.call(this, "UpdateDocument");
  }

}
