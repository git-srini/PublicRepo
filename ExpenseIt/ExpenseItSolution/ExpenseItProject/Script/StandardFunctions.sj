﻿
/**  
*@function Launch_ExpenseIt
*@Summary function to launch the application ExpenseIt
*@Parameters
*/

function Launch_ExpenseIt() {
    try {
        if (!Sys.WaitProcess("ExpenseIt9", 500).Exists) {
            TestedApps.ExpenseIt9.Run();
            Sys.Process("ExpenseIt9").WaitProperty("Name", "Process('ExpenseIt9')", 1000);
        }
    } catch (err) {
        Log.Error(err.message);
    }
}

/**  
*@function GetExpensesDockPanelObject
*@Summary function to get the object control for expenses dock panel
*@Parameters
*/
function GetExpensesDockPanelObject() {
    try {
        dockPanelObject = Aliases.ExpensesIt.FindEx(['ClrClassName', 'Visible'], ['DockPanel', 'true'], 6, true, 1000);
        if (dockPanelObject.Exists) {
            return dockPanelObject;
        } else {
            Log.Error("dockPanel Object not available");
        }
    } catch (err) {
        Log.Error(err.message)
    }
}

/**  
*@function ActionOnTextbox
*@Summary 
*/
function ActionOnTextbox(oControl, value) {
    oControl.SetText(value);
}

/**  
*@function ActionOnCheckbox
*@Summary
*/
function ActionOnCheckbox(oControl, Action) {
    try {
        if (Action.trim().toUpperCase() == 'CHECK' || Action.trim().toUpperCase() == 'SELECT') {
            if (oControl.wState == 0) {
                oControl.ClickButton(cbChecked);
            }
        } else if (Action.trim().toUpperCase() == 'UNCHECK' || Action.trim().toUpperCase() == 'UNSELECT') {
            if (oControl.wState == 1) {
                oControl.ClickButton(cbUnchecked);
            }
        }
    } catch (err) {
        Log.Error(err.message);
    }
}

/**  
*@function SelectComboBox
*@Summary 
*/
function SelectComboBox(oControl, value) {
    try {
        oControl.ClickItem(value);
    } catch (err) {
        Log.Error(err.message);
    }
}

/**  
*@function ActionOnButton
*@Summary 
*/
function ActionOnButton(oControl, Action) {
    try {
        if (Action.toUpperCase() == 'SELECT') {
            if (oControl.Pressed == false) {
                oControl.Click();
            }
        } else if (Action.toUpperCase() == 'UNSELECT') {
            if (oControl.Pressed == true) {
                oControl.Click();
            }
        } else if (Action.toUpperCase() == 'CLICK') {
            oControl.Click();
        } else if (Action.toUpperCase() == 'DBLCLICK') {
            oControl.DblClick();
        }
    } catch (err) {
        Log.Error(err.message);
    }
}