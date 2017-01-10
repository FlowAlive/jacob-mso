/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Assistant extends Dispatch {

  public static final String componentName = "Office.Assistant";

  public Assistant() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public Assistant(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public Assistant(String compName) {
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
   * @param xLeft an input-parameter of type int
   * @param yTop an input-parameter of type int
   */
  public void move(int xLeft, int yTop) {
    Dispatch.call(this, "Move", new Variant(xLeft), new Variant(yTop));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param top an input-parameter of type int
   */
  public void setTop(int top) {
    Dispatch.put(this, "Top", new Variant(top));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getTop() {
    return Dispatch.get(this, "Top").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param left an input-parameter of type int
   */
  public void setLeft(int left) {
    Dispatch.put(this, "Left", new Variant(left));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getLeft() {
    return Dispatch.get(this, "Left").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void help() {
    Dispatch.call(this, "Help");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param on an input-parameter of type boolean
   * @param callback an input-parameter of type String
   * @param privateX an input-parameter of type int
   * @param animation an input-parameter of type Variant
   * @param customTeaser an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param bottom an input-parameter of type Variant
   * @param right an input-parameter of type Variant
   * @return the result is of type int
   */
  public int startWizard(boolean on, String callback, int privateX, Variant animation, Variant customTeaser,
                         Variant top, Variant left, Variant bottom, Variant right) {
    return Dispatch.callN(this, "StartWizard", new Object[] {new Variant(on), callback, new Variant(privateX),
                          animation, customTeaser, top, left, bottom, right}).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param on an input-parameter of type boolean
   * @param callback an input-parameter of type String
   * @param privateX an input-parameter of type int
   * @param animation an input-parameter of type Variant
   * @param customTeaser an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @param bottom an input-parameter of type Variant
   * @return the result is of type int
   */
  public int startWizard(boolean on, String callback, int privateX, Variant animation, Variant customTeaser,
                         Variant top, Variant left, Variant bottom) {
    return Dispatch.call(this, "StartWizard", new Variant(on), callback, new Variant(privateX), animation, customTeaser,
                         top, left, bottom).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param on an input-parameter of type boolean
   * @param callback an input-parameter of type String
   * @param privateX an input-parameter of type int
   * @param animation an input-parameter of type Variant
   * @param customTeaser an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @param left an input-parameter of type Variant
   * @return the result is of type int
   */
  public int startWizard(boolean on, String callback, int privateX, Variant animation, Variant customTeaser,
                         Variant top, Variant left) {
    return Dispatch.call(this, "StartWizard", new Variant(on), callback, new Variant(privateX), animation, customTeaser,
                         top, left).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param on an input-parameter of type boolean
   * @param callback an input-parameter of type String
   * @param privateX an input-parameter of type int
   * @param animation an input-parameter of type Variant
   * @param customTeaser an input-parameter of type Variant
   * @param top an input-parameter of type Variant
   * @return the result is of type int
   */
  public int startWizard(boolean on, String callback, int privateX, Variant animation, Variant customTeaser,
                         Variant top) {
    return Dispatch.call(this, "StartWizard", new Variant(on), callback, new Variant(privateX), animation, customTeaser,
                         top).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param on an input-parameter of type boolean
   * @param callback an input-parameter of type String
   * @param privateX an input-parameter of type int
   * @param animation an input-parameter of type Variant
   * @param customTeaser an input-parameter of type Variant
   * @return the result is of type int
   */
  public int startWizard(boolean on, String callback, int privateX, Variant animation, Variant customTeaser) {
    return Dispatch.call(this, "StartWizard", new Variant(on), callback, new Variant(privateX), animation, customTeaser).
            changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param on an input-parameter of type boolean
   * @param callback an input-parameter of type String
   * @param privateX an input-parameter of type int
   * @param animation an input-parameter of type Variant
   * @return the result is of type int
   */
  public int startWizard(boolean on, String callback, int privateX, Variant animation) {
    return Dispatch.call(this, "StartWizard", new Variant(on), callback, new Variant(privateX),
            animation).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param on an input-parameter of type boolean
   * @param callback an input-parameter of type String
   * @param privateX an input-parameter of type int
   * @return the result is of type int
   */
  public int startWizard(boolean on, String callback, int privateX) {
    return Dispatch.call(this, "StartWizard", new Variant(on), callback,
            new Variant(privateX)).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param wizardID an input-parameter of type int
   * @param varfSuccess an input-parameter of type boolean
   * @param animation an input-parameter of type Variant
   */
  public void endWizard(int wizardID, boolean varfSuccess, Variant animation) {
    Dispatch.call(this, "EndWizard", new Variant(wizardID), new Variant(varfSuccess), animation);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param wizardID an input-parameter of type int
   * @param varfSuccess an input-parameter of type boolean
   */
  public void endWizard(int wizardID, boolean varfSuccess) {
    Dispatch.call(this, "EndWizard", new Variant(wizardID), new Variant(varfSuccess));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param wizardID an input-parameter of type int
   * @param act an input-parameter of type int
   * @param animation an input-parameter of type Variant
   */
  public void activateWizard(int wizardID, int act, Variant animation) {
    Dispatch.call(this, "ActivateWizard", new Variant(wizardID), new Variant(act), animation);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param wizardID an input-parameter of type int
   * @param act an input-parameter of type int
   */
  public void activateWizard(int wizardID, int act) {
    Dispatch.call(this, "ActivateWizard", new Variant(wizardID), new Variant(act));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void resetTips() {
    Dispatch.call(this, "ResetTips");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Balloon
   */
  public Balloon getNewBalloon() {
    return new Balloon(Dispatch.get(this, "NewBalloon").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getBalloonError() {
    return Dispatch.get(this, "BalloonError").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getVisible() {
    return Dispatch.get(this, "Visible").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param visible an input-parameter of type boolean
   */
  public void setVisible(boolean visible) {
    Dispatch.put(this, "Visible", new Variant(visible));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getAnimation() {
    return Dispatch.get(this, "Animation").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param animation an input-parameter of type int
   */
  public void setAnimation(int animation) {
    Dispatch.put(this, "Animation", new Variant(animation));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getReduced() {
    return Dispatch.get(this, "Reduced").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param reduced an input-parameter of type boolean
   */
  public void setReduced(boolean reduced) {
    Dispatch.put(this, "Reduced", new Variant(reduced));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param assistWithHelp an input-parameter of type boolean
   */
  public void setAssistWithHelp(boolean assistWithHelp) {
    Dispatch.put(this, "AssistWithHelp", new Variant(assistWithHelp));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getAssistWithHelp() {
    return Dispatch.get(this, "AssistWithHelp").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param assistWithWizards an input-parameter of type boolean
   */
  public void setAssistWithWizards(boolean assistWithWizards) {
    Dispatch.put(this, "AssistWithWizards", new Variant(assistWithWizards));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getAssistWithWizards() {
    return Dispatch.get(this, "AssistWithWizards").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param assistWithAlerts an input-parameter of type boolean
   */
  public void setAssistWithAlerts(boolean assistWithAlerts) {
    Dispatch.put(this, "AssistWithAlerts", new Variant(assistWithAlerts));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getAssistWithAlerts() {
    return Dispatch.get(this, "AssistWithAlerts").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param moveWhenInTheWay an input-parameter of type boolean
   */
  public void setMoveWhenInTheWay(boolean moveWhenInTheWay) {
    Dispatch.put(this, "MoveWhenInTheWay", new Variant(moveWhenInTheWay));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getMoveWhenInTheWay() {
    return Dispatch.get(this, "MoveWhenInTheWay").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param sounds an input-parameter of type boolean
   */
  public void setSounds(boolean sounds) {
    Dispatch.put(this, "Sounds", new Variant(sounds));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getSounds() {
    return Dispatch.get(this, "Sounds").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param featureTips an input-parameter of type boolean
   */
  public void setFeatureTips(boolean featureTips) {
    Dispatch.put(this, "FeatureTips", new Variant(featureTips));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getFeatureTips() {
    return Dispatch.get(this, "FeatureTips").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param mouseTips an input-parameter of type boolean
   */
  public void setMouseTips(boolean mouseTips) {
    Dispatch.put(this, "MouseTips", new Variant(mouseTips));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getMouseTips() {
    return Dispatch.get(this, "MouseTips").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param keyboardShortcutTips an input-parameter of type boolean
   */
  public void setKeyboardShortcutTips(boolean keyboardShortcutTips) {
    Dispatch.put(this, "KeyboardShortcutTips", new Variant(keyboardShortcutTips));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getKeyboardShortcutTips() {
    return Dispatch.get(this, "KeyboardShortcutTips").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param highPriorityTips an input-parameter of type boolean
   */
  public void setHighPriorityTips(boolean highPriorityTips) {
    Dispatch.put(this, "HighPriorityTips", new Variant(highPriorityTips));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getHighPriorityTips() {
    return Dispatch.get(this, "HighPriorityTips").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param tipOfDay an input-parameter of type boolean
   */
  public void setTipOfDay(boolean tipOfDay) {
    Dispatch.put(this, "TipOfDay", new Variant(tipOfDay));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getTipOfDay() {
    return Dispatch.get(this, "TipOfDay").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param guessHelp an input-parameter of type boolean
   */
  public void setGuessHelp(boolean guessHelp) {
    Dispatch.put(this, "GuessHelp", new Variant(guessHelp));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getGuessHelp() {
    return Dispatch.get(this, "GuessHelp").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param searchWhenProgramming an input-parameter of type boolean
   */
  public void setSearchWhenProgramming(boolean searchWhenProgramming) {
    Dispatch.put(this, "SearchWhenProgramming", new Variant(searchWhenProgramming));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getSearchWhenProgramming() {
    return Dispatch.get(this, "SearchWhenProgramming").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getItem() {
    return Dispatch.get(this, "Item").toString();
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
   * @return the result is of type String
   */
  public String getName() {
    return Dispatch.get(this, "Name").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getOn() {
    return Dispatch.get(this, "On").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param on an input-parameter of type boolean
   */
  public void setOn(boolean on) {
    Dispatch.put(this, "On", new Variant(on));
  }

}
