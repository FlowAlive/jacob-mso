/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.word;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class EmailOptions extends Dispatch {

  public static final String componentName = "Word.EmailOptions";

  public EmailOptions() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public EmailOptions(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public EmailOptions(String compName) {
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
   * @return the result is of type boolean
   */
  public boolean getUseThemeStyle() {
    return Dispatch.get(this, "UseThemeStyle").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param useThemeStyle an input-parameter of type boolean
   */
  public void setUseThemeStyle(boolean useThemeStyle) {
    Dispatch.put(this, "UseThemeStyle", new Variant(useThemeStyle));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getMarkCommentsWith() {
    return Dispatch.get(this, "MarkCommentsWith").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param markCommentsWith an input-parameter of type String
   */
  public void setMarkCommentsWith(String markCommentsWith) {
    Dispatch.put(this, "MarkCommentsWith", markCommentsWith);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getMarkComments() {
    return Dispatch.get(this, "MarkComments").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param markComments an input-parameter of type boolean
   */
  public void setMarkComments(boolean markComments) {
    Dispatch.put(this, "MarkComments", new Variant(markComments));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type EmailSignature
   */
  public EmailSignature getEmailSignature() {
    return new EmailSignature(Dispatch.get(this, "EmailSignature").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Style
   */
  public Style getComposeStyle() {
    return new Style(Dispatch.get(this, "ComposeStyle").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Style
   */
  public Style getReplyStyle() {
    return new Style(Dispatch.get(this, "ReplyStyle").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getThemeName() {
    return Dispatch.get(this, "ThemeName").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param themeName an input-parameter of type String
   */
  public void setThemeName(String themeName) {
    Dispatch.put(this, "ThemeName", themeName);
  }

}
