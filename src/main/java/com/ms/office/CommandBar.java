/**
 * JacobGen generated file --- do not edit
 *
 * (http://www.sourceforge.net/projects/jacob-project */
package com.ms.office;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class CommandBar extends Dispatch {

  public static final String componentName = "Office.CommandBar";

  public CommandBar() {
    super(componentName);
  }

  /**
   * This constructor is used instead of a case operation to
   * turn a Dispatch object into a wider object - it must exist
   * in every wrapper class whose instances may be returned from
   * method calls wrapped in VT_DISPATCH Variants.
   */
  public CommandBar(Dispatch d) {
    // take over the IDispatch pointer
    m_pDispatch = d.m_pDispatch;
    // null out the input's pointer
    d.m_pDispatch = 0;
  }

  public CommandBar(String compName) {
    super(compName);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Object
   */
  public Object getaccParent() {
    return Dispatch.get(this, "accParent");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getaccChildCount() {
    return Dispatch.get(this, "accChildCount").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   * @return the result is of type Object
   */
  public Object getaccChild(Variant varChild) {
    return Dispatch.call(this, "accChild", varChild);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   * @return the result is of type String
   */
  public String getaccName(Variant varChild) {
    return Dispatch.call(this, "accName", varChild).toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getaccName() {
    return Dispatch.get(this, "accName").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   * @return the result is of type String
   */
  public String getaccValue(Variant varChild) {
    return Dispatch.call(this, "accValue", varChild).toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getaccValue() {
    return Dispatch.get(this, "accValue").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   * @return the result is of type String
   */
  public String getaccDescription(Variant varChild) {
    return Dispatch.call(this, "accDescription", varChild).toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getaccDescription() {
    return Dispatch.get(this, "accDescription").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   * @return the result is of type Variant
   */
  public Variant getaccRole(Variant varChild) {
    return Dispatch.call(this, "accRole", varChild);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant getaccRole() {
    return Dispatch.get(this, "accRole");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   * @return the result is of type Variant
   */
  public Variant getaccState(Variant varChild) {
    return Dispatch.call(this, "accState", varChild);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant getaccState() {
    return Dispatch.get(this, "accState");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   * @return the result is of type String
   */
  public String getaccHelp(Variant varChild) {
    return Dispatch.call(this, "accHelp", varChild).toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getaccHelp() {
    return Dispatch.get(this, "accHelp").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param pszHelpFile an input-parameter of type String
   * @param varChild an input-parameter of type Variant
   * @return the result is of type int
   */
  public int getaccHelpTopic(String pszHelpFile, Variant varChild) {
    return Dispatch.call(this, "accHelpTopic", pszHelpFile, varChild).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param pszHelpFile an input-parameter of type String
   * @return the result is of type int
   */
  public int getaccHelpTopic(String pszHelpFile) {
    return Dispatch.call(this, "accHelpTopic", pszHelpFile).changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
   * @param pszHelpFile is an one-element array which sends the input-parameter
   *                    to the ActiveX-Component and receives the output-parameter
   * @param varChild an input-parameter of type Variant
   * @return the result is of type int
   */
  public int getaccHelpTopic(String[] pszHelpFile, Variant varChild) {
    Variant vnt_pszHelpFile = new Variant();
    if (pszHelpFile == null || pszHelpFile.length == 0) {
      vnt_pszHelpFile.putNoParam();
    }
    else {
      vnt_pszHelpFile.putStringRef(pszHelpFile[0]);
    }

    int result_of_accHelpTopic = Dispatch.call(this, "accHelpTopic", vnt_pszHelpFile,
                                 varChild).changeType(Variant.VariantInt).getInt();

    if (pszHelpFile != null && pszHelpFile.length > 0) {
      pszHelpFile[0] = vnt_pszHelpFile.toString();
    }

    return result_of_accHelpTopic;
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   * @return the result is of type String
   */
  public String getaccKeyboardShortcut(Variant varChild) {
    return Dispatch.call(this, "accKeyboardShortcut", varChild).toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getaccKeyboardShortcut() {
    return Dispatch.get(this, "accKeyboardShortcut").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant getaccFocus() {
    return Dispatch.get(this, "accFocus");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type Variant
   */
  public Variant getaccSelection() {
    return Dispatch.get(this, "accSelection");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   * @return the result is of type String
   */
  public String getaccDefaultAction(Variant varChild) {
    return Dispatch.call(this, "accDefaultAction", varChild).toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getaccDefaultAction() {
    return Dispatch.get(this, "accDefaultAction").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param flagsSelect an input-parameter of type int
   * @param varChild an input-parameter of type Variant
   */
  public void accSelect(int flagsSelect, Variant varChild) {
    Dispatch.call(this, "accSelect", new Variant(flagsSelect), varChild);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param flagsSelect an input-parameter of type int
   */
  public void accSelect(int flagsSelect) {
    Dispatch.call(this, "accSelect", new Variant(flagsSelect));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param pxLeft an input-parameter of type int
   * @param pyTop an input-parameter of type int
   * @param pcxWidth an input-parameter of type int
   * @param pcyHeight an input-parameter of type int
   * @param varChild an input-parameter of type Variant
   */
  public void accLocation(int pxLeft, int pyTop, int pcxWidth, int pcyHeight, Variant varChild) {
    Dispatch.call(this, "accLocation", new Variant(pxLeft), new Variant(pyTop), new Variant(pcxWidth),
                  new Variant(pcyHeight), varChild);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param pxLeft an input-parameter of type int
   * @param pyTop an input-parameter of type int
   * @param pcxWidth an input-parameter of type int
   * @param pcyHeight an input-parameter of type int
   */
  public void accLocation(int pxLeft, int pyTop, int pcxWidth, int pcyHeight) {
    Dispatch.call(this, "accLocation", new Variant(pxLeft), new Variant(pyTop), new Variant(pcxWidth),
                  new Variant(pcyHeight));
  }

  /**
   * Wrapper for calling the ActiveX-Method and receiving the output-parameter(s).
   * @param pxLeft is an one-element array which sends the input-parameter
   *               to the ActiveX-Component and receives the output-parameter
   * @param pyTop is an one-element array which sends the input-parameter
   *              to the ActiveX-Component and receives the output-parameter
   * @param pcxWidth is an one-element array which sends the input-parameter
   *                 to the ActiveX-Component and receives the output-parameter
   * @param pcyHeight is an one-element array which sends the input-parameter
   *                  to the ActiveX-Component and receives the output-parameter
   * @param varChild an input-parameter of type Variant
   */
  public void accLocation(int[] pxLeft, int[] pyTop, int[] pcxWidth, int[] pcyHeight, Variant varChild) {
    Variant vnt_pxLeft = new Variant();
    if (pxLeft == null || pxLeft.length == 0) {
      vnt_pxLeft.putNoParam();
    }
    else {
      vnt_pxLeft.putIntRef(pxLeft[0]);
    }

    Variant vnt_pyTop = new Variant();
    if (pyTop == null || pyTop.length == 0) {
      vnt_pyTop.putNoParam();
    }
    else {
      vnt_pyTop.putIntRef(pyTop[0]);
    }

    Variant vnt_pcxWidth = new Variant();
    if (pcxWidth == null || pcxWidth.length == 0) {
      vnt_pcxWidth.putNoParam();
    }
    else {
      vnt_pcxWidth.putIntRef(pcxWidth[0]);
    }

    Variant vnt_pcyHeight = new Variant();
    if (pcyHeight == null || pcyHeight.length == 0) {
      vnt_pcyHeight.putNoParam();
    }
    else {
      vnt_pcyHeight.putIntRef(pcyHeight[0]);
    }

    Dispatch.call(this, "accLocation", vnt_pxLeft, vnt_pyTop, vnt_pcxWidth, vnt_pcyHeight, varChild);

    if (pxLeft != null && pxLeft.length > 0) {
      pxLeft[0] = vnt_pxLeft.toInt();
    }
    if (pyTop != null && pyTop.length > 0) {
      pyTop[0] = vnt_pyTop.toInt();
    }
    if (pcxWidth != null && pcxWidth.length > 0) {
      pcxWidth[0] = vnt_pcxWidth.toInt();
    }
    if (pcyHeight != null && pcyHeight.length > 0) {
      pcyHeight[0] = vnt_pcyHeight.toInt();
    }
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param navDir an input-parameter of type int
   * @param varStart an input-parameter of type Variant
   * @return the result is of type Variant
   */
  public Variant accNavigate(int navDir, Variant varStart) {
    return Dispatch.call(this, "accNavigate", new Variant(navDir), varStart);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param navDir an input-parameter of type int
   * @return the result is of type Variant
   */
  public Variant accNavigate(int navDir) {
    return Dispatch.call(this, "accNavigate", new Variant(navDir));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param xLeft an input-parameter of type int
   * @param yTop an input-parameter of type int
   * @return the result is of type Variant
   */
  public Variant accHitTest(int xLeft, int yTop) {
    return Dispatch.call(this, "accHitTest", new Variant(xLeft), new Variant(yTop));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   */
  public void accDoDefaultAction(Variant varChild) {
    Dispatch.call(this, "accDoDefaultAction", varChild);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void accDoDefaultAction() {
    Dispatch.call(this, "accDoDefaultAction");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   */
  public void setaccName(Variant varChild) {
    Dispatch.put(this, "accName", varChild);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void setaccName() {
    Dispatch.call(this, "accName");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param varChild an input-parameter of type Variant
   */
  public void setaccValue(Variant varChild) {
    Dispatch.put(this, "accValue", varChild);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void setaccValue() {
    Dispatch.call(this, "accValue");
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
  public boolean getBuiltIn() {
    return Dispatch.get(this, "BuiltIn").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getContext() {
    return Dispatch.get(this, "Context").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param context an input-parameter of type String
   */
  public void setContext(String context) {
    Dispatch.put(this, "Context", context);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void delete() {
    Dispatch.call(this, "Delete");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getEnabled() {
    return Dispatch.get(this, "Enabled").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param enabled an input-parameter of type boolean
   */
  public void setEnabled(boolean enabled) {
    Dispatch.put(this, "Enabled", new Variant(enabled));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param type an input-parameter of type Variant
   * @param id an input-parameter of type Variant
   * @param tag an input-parameter of type Variant
   * @param visible an input-parameter of type Variant
   * @param recursive an input-parameter of type Variant
   * @return the result is of type CommandBarControl
   */
  public CommandBarControl findControl(Variant type, Variant id, Variant tag, Variant visible, Variant recursive) {
    return new CommandBarControl(Dispatch.call(this, "FindControl", type, id, tag, visible, recursive).toDispatch());
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
   * @return the result is of type int
   */
  public int getHeight() {
    return Dispatch.get(this, "Height").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param height an input-parameter of type int
   */
  public void setHeight(int height) {
    Dispatch.put(this, "Height", new Variant(height));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getIndex() {
    return Dispatch.get(this, "Index").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getInstanceId() {
    return Dispatch.get(this, "InstanceId").changeType(Variant.VariantInt).getInt();
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
   * @param left an input-parameter of type int
   */
  public void setLeft(int left) {
    Dispatch.put(this, "Left", new Variant(left));
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
   * @param name an input-parameter of type String
   */
  public void setName(String name) {
    Dispatch.put(this, "Name", name);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type String
   */
  public String getNameLocal() {
    return Dispatch.get(this, "NameLocal").toString();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param nameLocal an input-parameter of type String
   */
  public void setNameLocal(String nameLocal) {
    Dispatch.put(this, "NameLocal", nameLocal);
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
  public int getPosition() {
    return Dispatch.get(this, "Position").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param position an input-parameter of type int
   */
  public void setPosition(int position) {
    Dispatch.put(this, "Position", new Variant(position));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getRowIndex() {
    return Dispatch.get(this, "RowIndex").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param rowIndex an input-parameter of type int
   */
  public void setRowIndex(int rowIndex) {
    Dispatch.put(this, "RowIndex", new Variant(rowIndex));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type int
   */
  public int getProtection() {
    return Dispatch.get(this, "Protection").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param protection an input-parameter of type int
   */
  public void setProtection(int protection) {
    Dispatch.put(this, "Protection", new Variant(protection));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void reset() {
    Dispatch.call(this, "Reset");
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param x an input-parameter of type Variant
   * @param y an input-parameter of type Variant
   */
  public void showPopup(Variant x, Variant y) {
    Dispatch.call(this, "ShowPopup", x, y);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param x an input-parameter of type Variant
   */
  public void showPopup(Variant x) {
    Dispatch.call(this, "ShowPopup", x);
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   */
  public void showPopup() {
    Dispatch.call(this, "ShowPopup");
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
   * @param top an input-parameter of type int
   */
  public void setTop(int top) {
    Dispatch.put(this, "Top", new Variant(top));
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
  public int getWidth() {
    return Dispatch.get(this, "Width").changeType(Variant.VariantInt).getInt();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param width an input-parameter of type int
   */
  public void setWidth(int width) {
    Dispatch.put(this, "Width", new Variant(width));
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @return the result is of type boolean
   */
  public boolean getAdaptiveMenu() {
    return Dispatch.get(this, "AdaptiveMenu").changeType(Variant.VariantBoolean).getBoolean();
  }

  /**
   * Wrapper for calling the ActiveX-Method with input-parameter(s).
   * @param adaptiveMenu an input-parameter of type boolean
   */
  public void setAdaptiveMenu(boolean adaptiveMenu) {
    Dispatch.put(this, "AdaptiveMenu", new Variant(adaptiveMenu));
  }

}
