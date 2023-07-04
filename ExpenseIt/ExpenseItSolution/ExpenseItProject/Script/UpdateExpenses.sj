//USEUNIT StandardFunctions

/**  
*@function Update_Expenses
*@Summary function to Update Expenses
*/
function Update_Expenses(expType, expDescription, expCost, setType, setDescription, setCost, Action) {
    try {
        var expenseDataGridObject = Aliases.CreateExpenseReportDialog.FindEx(['ClrClassName', 'WPFControlName', 'Visible'], ['DataGrid', 'expenseDataGrid1', true], 8, true, 1000);
        if (expenseDataGridObject.Exists) {
            itemsCount = expenseDataGridObject.Items.Count
            for (i = 0; i < itemsCount; i++) {
                expItem = expenseDataGridObject.Items.item(i)
                switch (Action.toUpperCase()) {
                    case 'TYPE':
                        if (expItem.Type == expType) {
                            expItem.set_Type(setType);
                        }
                        break;

                    case 'DESCRIPTION':
                        if (expItem.Description == expDescription) {
                            expItem.set_Description(setDescription);
                        }
                        break;

                    case 'COST':
                        if (expItem.Cost == expCost) {
                            expItem.set_Cost(setCost);
                        }
                        break;

                    default:
                        Log.Error('Invalid Keyword-${err.message}');
                }
            }
        }
    } catch (err) {
        Log.Error(err.message);
    }
}

/**  
*@function UpdateTheExpenseReport
*@Summary function to Update the Expense Report
*/
function UpdateTheExpenseReport() {
    try {
        Launch_ExpenseIt();
        var dockPanelObject = GetExpensesDockPanelObject();
        if (dockPanelObject.Exists) {
            var createExpensesObject = dockPanelObject.FindEx(['ClrClassName', 'WPFControlText', 'Visible'], ['Button', 'Create Expense _Report', 'true'], 50, true, 1000);
            if (createExpensesObject.Exists) {
                ActionOnButton(createExpensesObject, 'Click');
            }
        }
        Update_Expenses("Private_Travel", "Bus_Cab", 2000, "Private Travel234", "Bus Cab234", 5000, "Type");
        Update_Expenses("Private_Travel", "Bus_Cab", 2000, "Private Travel234", "Bus Cab234", 5000, "Description");
        Update_Expenses("Private_Travel", "Bus_Cab", 2000, "Private Travel234", "Bus Cab234", 5000, "Cost");
    } catch (err) {
        Log.Error(err.message);
    }
}