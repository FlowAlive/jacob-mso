/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class _CommandBars extends Dispatch {

  public static final String componentName = "Office._CommandBars";

  public _CommandBars() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public _CommandBars(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public _CommandBars(String compName) {
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
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl getActionControl() {
    return new CommandBarControl(Dispatch.get(this, "ActionControl").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type CommandBar
   */
  public CommandBar getActiveMenuBar() {
    return new CommandBar(Dispatch.get(this, "ActiveMenuBar").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param name an input-parameter of type Variant
   * @param position an input-parameter of type Variant
   * @param menuBar an input-parameter of type Variant
   * @param temporary an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar add(Variant name, Variant position, Variant menuBar, Variant temporary) {
    return new CommandBar(Dispatch.call(this, "Add", name, position, menuBar, temporary).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param name an input-parameter of type Variant
   * @param position an input-parameter of type Variant
   * @param menuBar an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar add(Variant name, Variant position, Variant menuBar) {
    return new CommandBar(Dispatch.call(this, "Add", name, position, menuBar).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param name an input-parameter of type Variant
   * @param position an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar add(Variant name, Variant position) {
    return new CommandBar(Dispatch.call(this, "Add", name, position).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param name an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar add(Variant name) {
    return new CommandBar(Dispatch.call(this, "Add", name).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type CommandBar
   */
  public CommandBar add() {
    return new CommandBar(Dispatch.call(this, "Add").toDispatch());
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
   * @return the result is of type boolean
   */
  public boolean getDisplayTooltips() {
    return Dispatch.get(this, "DisplayTooltips").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param displayTooltips an input-parameter of type boolean
   */
  public void setDisplayTooltips(boolean displayTooltips) {
    Dispatch.put(this, "DisplayTooltips", new Variant(displayTooltips));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getDisplayKeysInTooltips() {
    return Dispatch.get(this, "DisplayKeysInTooltips").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param displayKeysInTooltips an input-parameter of type boolean
   */
  public void setDisplayKeysInTooltips(boolean displayKeysInTooltips) {
    Dispatch.put(this, "DisplayKeysInTooltips", new Variant(displayKeysInTooltips));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @param tag an input-parameter of type Variant
   * @param visible an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl findControl(Variant type, Variant id, Variant tag, Variant visible) {
    return new CommandBarControl(Dispatch.call(this, "FindControl", type, id, tag, visible).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @param tag an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl findControl(Variant type, Variant id, Variant tag) {
    return new CommandBarControl(Dispatch.call(this, "FindControl", type, id, tag).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl findControl(Variant type, Variant id) {
    return new CommandBarControl(Dispatch.call(this, "FindControl", type, id).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl findControl(Variant type) {
    return new CommandBarControl(Dispatch.call(this, "FindControl", type).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl findControl() {
    return new CommandBarControl(Dispatch.call(this, "FindControl").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param index an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar getItem(Variant index) {
    return new CommandBar(Dispatch.call(this, "Item", index).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getLargeButtons() {
    return Dispatch.get(this, "LargeButtons").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param largeButtons an input-parameter of type boolean
   */
  public void setLargeButtons(boolean largeButtons) {
    Dispatch.put(this, "LargeButtons", new Variant(largeButtons));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getMenuAnimationStyle() {
    return Dispatch.get(this, "MenuAnimationStyle").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param menuAnimationStyle an input-parameter of type int
   */
  public void setMenuAnimationStyle(int menuAnimationStyle) {
    Dispatch.put(this, "MenuAnimationStyle", new Variant(menuAnimationStyle));
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
   * @return the result is of type Object
   */
  public Object getParent() {
    return Dispatch.get(this, "Parent");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void releaseFocus() {
    Dispatch.call(this, "ReleaseFocus");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param ids an input-parameter of type int
   * @param pbstrName an input-parameter of type String
   * @return the result is of type int
   */
  public int getIdsString(int ids, String pbstrName) {
    return Dispatch.call(this, "IdsString", new Variant(ids), pbstrName).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
   * @param ids an input-parameter of type int
   * @param pbstrName is an one-element array which sends the input-parameter
   *                  to the ActiveX-Component and receives the output-parameter
   * @return the result is of type int
   */
  public int getIdsString(int ids, String[] pbstrName) {
    Variant vnt_pbstrName = new Variant();
    if (pbstrName == null || pbstrName.length == 0) {
      vnt_pbstrName.putNoParam();
    }
    else {
      vnt_pbstrName.putStringRef(pbstrName[0]);
    }

    int result_of_IdsString = Dispatch.call(this, "IdsString", new Variant(ids),
                              vnt_pbstrName).changeType(Variant.VariantInt).getInt();

    if (pbstrName != null && pbstrName.length > 0) {
      pbstrName[0] = vnt_pbstrName.toString();
    }

    return result_of_IdsString;
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param tmc an input-parameter of type int
   * @param pbstrName an input-parameter of type String
   * @return the result is of type int
   */
  public int getTmcGetName(int tmc, String pbstrName) {
    return Dispatch.call(this, "TmcGetName", new Variant(tmc), pbstrName).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
   * @param tmc an input-parameter of type int
   * @param pbstrName is an one-element array which sends the input-parameter
   *                  to the ActiveX-Component and receives the output-parameter
   * @return the result is of type int
   */
  public int getTmcGetName(int tmc, String[] pbstrName) {
    Variant vnt_pbstrName = new Variant();
    if (pbstrName == null || pbstrName.length == 0) {
      vnt_pbstrName.putNoParam();
    }
    else {
      vnt_pbstrName.putStringRef(pbstrName[0]);
    }

    int result_of_TmcGetName = Dispatch.call(this, "TmcGetName", new Variant(tmc),
                               vnt_pbstrName).changeType(Variant.VariantInt).getInt();

    if (pbstrName != null && pbstrName.length > 0) {
      pbstrName[0] = vnt_pbstrName.toString();
    }

    return result_of_TmcGetName;
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getAdaptiveMenus() {
    return Dispatch.get(this, "AdaptiveMenus").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param adaptiveMenus an input-parameter of type boolean
   */
  public void setAdaptiveMenus(boolean adaptiveMenus) {
    Dispatch.put(this, "AdaptiveMenus", new Variant(adaptiveMenus));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @param tag an input-parameter of type Variant
   * @param visible an input-parameter of type Variant
   * @return the result is of type CommandBarControls
   */
  public CommandBarControls findControls(Variant type, Variant id, Variant tag, Variant visible) {
    return new CommandBarControls(Dispatch.call(this, "FindControls", type, id, tag, visible).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @param tag an input-parameter of type Variant
   * @return the result is of type CommandBarControls
   */
  public CommandBarControls findControls(Variant type, Variant id, Variant tag) {
    return new CommandBarControls(Dispatch.call(this, "FindControls", type, id, tag).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @return the result is of type CommandBarControls
   */
  public CommandBarControls findControls(Variant type, Variant id) {
    return new CommandBarControls(Dispatch.call(this, "FindControls", type, id).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @return the result is of type CommandBarControls
   */
  public CommandBarControls findControls(Variant type) {
    return new CommandBarControls(Dispatch.call(this, "FindControls", type).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type CommandBarControls
   */
  public CommandBarControls findControls() {
    return new CommandBarControls(Dispatch.call(this, "FindControls").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param tbidOrName an input-parameter of type Variant
   * @param position an input-parameter of type Variant
   * @param menuBar an input-parameter of type Variant
   * @param temporary an input-parameter of type Variant
   * @param tbtrProtection an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar addEx(Variant tbidOrName, Variant position, Variant menuBar, Variant temporary,
                          Variant tbtrProtection) {
    return new CommandBar(Dispatch.call(this, "AddEx", tbidOrName, position, menuBar, temporary, tbtrProtection).
                          toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param tbidOrName an input-parameter of type Variant
   * @param position an input-parameter of type Variant
   * @param menuBar an input-parameter of type Variant
   * @param temporary an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar addEx(Variant tbidOrName, Variant position, Variant menuBar, Variant temporary) {
    return new CommandBar(Dispatch.call(this, "AddEx", tbidOrName, position, menuBar, temporary).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param tbidOrName an input-parameter of type Variant
   * @param position an input-parameter of type Variant
   * @param menuBar an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar addEx(Variant tbidOrName, Variant position, Variant menuBar) {
    return new CommandBar(Dispatch.call(this, "AddEx", tbidOrName, position, menuBar).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param tbidOrName an input-parameter of type Variant
   * @param position an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar addEx(Variant tbidOrName, Variant position) {
    return new CommandBar(Dispatch.call(this, "AddEx", tbidOrName, position).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param tbidOrName an input-parameter of type Variant
   * @return the result is of type CommandBar
   */
  public CommandBar addEx(Variant tbidOrName) {
    return new CommandBar(Dispatch.call(this, "AddEx", tbidOrName).toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type CommandBar
   */
  public CommandBar addEx() {
    return new CommandBar(Dispatch.call(this, "AddEx").toDispatch());
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getDisplayFonts() {
    return Dispatch.get(this, "DisplayFonts").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param displayFonts an input-parameter of type boolean
   */
  public void setDisplayFonts(boolean displayFonts) {
    Dispatch.put(this, "DisplayFonts", new Variant(displayFonts));
  }

}
