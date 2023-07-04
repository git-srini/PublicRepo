//USEUNIT DeleteExpenses
//USEUNIT StandardFunctions
//USEUNIT UpdateExpenses


/**  
*@function Select_CostCenter
*@Summary function to select the cost center
*/
function Select_CostCenter(costCenter) {
    try {
        var costCenterObject = dockPanelObject.FindEx(['ClrClassName', 'WPFControlName', 'Visible'], ['ComboBox', 'costCenterTextBox', 'true'], 6, true, 1000);
        if (costCenterObject.Exists) {
            ActionOnButton(costCenterObject, "Click");
            SelectComboBox(costCenterObject, costCenter);
            Log.Checkpoint("Cost Center Item " + costCenter + " selected");
        } else {
            Log.Error("Cost Center Item object not available");
        }
    } catch (err) {
        Log.Error(err.message)
    }
}

/**  
*@function Set_EmployeeNumber
*@Summary function to set the Employee Number
*/
function Set_EmployeeNumber(employeeNumber) {
    try {
        var empNumberObject = dockPanelObject.FindEx(['ClrClassName', 'WPFControlName', 'Visible'], ['TextBox', 'employeeNumberTextBox', 'true'], 6, true, 1000);
        if (empNumberObject.Exists) {
            ActionOnTextbox(empNumberObject, employeeNumber);
            Log.Checkpoint("Employee number set as " + employeeNumber);
        } else {
            Log.Error("Employee number object not available");
        }
    } catch (err) {
        Log.Error(err.message)
    }
}

/**  
*@function Set_Email
*@Summary function to set the email id
*/
function Set_Email(emailId) {
    try {
        var emailIdObject = dockPanelObject.FindEx(['ClrClassName', 'WPFControlName', 'Visible'], ['TextBox', 'emailTextBox', 'true'], 6, true, 1000);
        if (emailIdObject.Exists) {
            ActionOnTextbox(emailIdObject, emailId);
            Log.Checkpoint("Email id set as " + emailId);
        } else {
            Log.Error("Email id object not available");
        }
    } catch (err) {
        Log.Error(err.message)
    }
}

/**  
*@function Select_EmployeeType
*@Summary function to select the Employee type
*/
function Select_EmployeeType(employeeType) {
    try {
        var employeeTypeObject = dockPanelObject.FindEx(['ClrClassName', 'WPFControlText', 'Visible'], ['ListBoxItem', employeeType, 'true'], 50, true, 1000);
        if (employeeTypeObject.Exists) {
            ActionOnButton(employeeTypeObject, "Click");
            Log.Checkpoint("Employee Type set as " + employeeType);
        } else {
            Log.Error("Employee Type object not available");
        }
    } catch (err) {
        Log.Error(err.message)
    }
}

/**Click_Add_ExpenseButton
*@function Click_Add_Expenses
*@Summary function to Click on Add Expense button
*/
function Click_Add_ExpenseButton() {
    try {
        if (Aliases.CreateExpenseReportDialog.Exists) {
            var addExpenseObject = Aliases.CreateExpenseReportDialog.FindEx(['ClrClassName', 'WPFControlText', 'Visible'], ['Button', 'Add _Expense', true], 8, true, 1000);
            if (addExpenseObject.Exists) {
                ActionOnButton(addExpenseObject, 'Click');
            }
        }
    } catch (err) {
        Log.Error(err.message);
    }
}

/**  
*@function Add_NewExpenses
*@Summary function to Add new Expenses
*/
function Add_NewExpenses(expType, expDescription, expCost) {
    try {
        var expenseDataGridObject = Aliases.CreateExpenseReportDialog.FindEx(['ClrClassName', 'WPFControlName', 'Visible'], ['DataGrid', 'expenseDataGrid1', true], 8, true, 1000);
        if (expenseDataGridObject.Exists) {
            itemsCount = expenseDataGridObject.Items.Count
            for (i = 0; i < itemsCount; i++) {
                expItem = expenseDataGridObject.Items.item(i)
                if (expItem.Type == "(Expense type)") {
                    expItem.set_Type(expType);
                }
                if (expItem.Description == "(Description)") {
                    expItem.set_Description(expDescription);
                }
                if (expItem.Cost == 0) {
                    expItem.set_Cost(expCost);
                }
            }
        }
        var OkButtonObject = Aliases.CreateExpenseReportDialog.FindEx(['ClrClassName', 'WPFControlText', 'Visible'], ['Button', '_OK', true], 8, true, 1000);
        if (OkButtonObject.Exists) {
            ActionOnButton(OkButtonObject, 'Click');
            if (Sys.Process("ExpenseIt9").Window("#32770", "ExpenseIt Standalone", 1).Window("Button", "OK", 1).Exists) {
                Sys.Process("ExpenseIt9").Window("#32770", "ExpenseIt Standalone", 1).Window("Button", "OK", 1).ClickButton();
            }
        }
    } catch (err) {
        Log.Error(err.message);
    }
}

/**  
*@function Create_ExpensesReport
*@Summary function to Create Expenses report
*/
function Create_ExpensesReport() {
    try {
        Launch_ExpenseIt();
        var dockPanelObject = GetExpensesDockPanelObject();
        if (dockPanelObject.Exists) {
            Set_Email("Srini@gmail.com");
            Set_EmployeeNumber("1234");
            Select_CostCenter(1); //1 -- for marketting // name is not working
            Select_EmployeeType("CSG");
            var createExpensesObject = dockPanelObject.FindEx(['ClrClassName', 'WPFControlText', 'Visible'], ['Button', 'Create Expense _Report', 'true'], 50, true, 1000);
            if (createExpensesObject.Exists) {
                ActionOnButton(createExpensesObject, 'Click');               
            }
            Click_Add_ExpenseButton();
            Add_NewExpenses("Private_Travel", "Bus_Cab", 2000);
        }
    } catch (err) {
        Log.Error(err.message);
    }
}