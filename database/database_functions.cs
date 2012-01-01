using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Collections;
using System.IO;
using System.Diagnostics;
using ADOX;
//using System.Windows.Forms;

namespace database
{
    public class database_functions
    {
        string pstrDB = @"c:\test.mdb";           // string pointing to database location

        private static void CreateAllTables()
        {
            string pstrDB = @"c:\test.mdb";

            //  CUSTOMERS table
            string[] customersColumns = { "P_Id", "LastName", "FirstName", "Address1", "Address2", "City", "State", "Zip", "ContactFirst", "ContactLast", "Fax", "CustomerID", "CreditLimit", "PriceLevel", "Phone1", "Phone2", "Contract", "StatementType", "ServiceCharge", "TaxExempt", "Dunning", "AcctClass", "LastOrderDate", "LastOrderAmount", "LastPaymentDate", "LastPaymentAmount", "OriginDate" };
            string[] customersColType = { "AUTOINCREMENT", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "int", "currency", "int", "text", "text", "yesno", "text", "yesno", "text", "yesno", "int", "Datetime", "currency", "Datetime", "currency", "datetime" };
            //  EMPLOYEE table
            string[] employeeColumns = { "P_Id", "EmployeeID", "EmployeeName", "PositionTitle", "OriginDate" };
            string[] employeeColType = { "AUTOINCREMENT", "int", "text", "text", "datetime" };
            //  PRODUCTS table
            string[] productsColumns = { "P_Id", "Pc", "Sku", "Subcat", "Description", "UnitOfMeasure", "Vendor", "VendorSku", "Manufacturer", "ManufacturerSku", "Upc", "AverageCost", "CurrentCost", "PriceLevel1", "PriceLevel2", "PriceLevel3", "PriceLevel4", "PriceLevel5", "PriceLevel6", "PriceLevel7", "PriceLevel8", "PriceLevel9", "PriceLevel10", "LeadIn", "MinOrderQty", "MinStockLvl", "MaxStockLvl", "OnHand" };
            string[] productsColType = { "AUTOINCREMENT", "int", "int", "int", "text", "text", "int", "text", "int", "text", "text", "Currency", "Currency", "Currency", "Currency", "Currency", "Currency", "Currency", "Currency", "Currency", "Currency", "Currency", "Currency", "int", "int", "int", "int", "int" };
            //  InventoryTransactions table
            string[] inventoryTransactionsColumns = { "P_Id", "Pc", "Sku", "Description", "Type", "TransactionType", "Invoice", "SaleDate", "Qty", "Before", "After" };
            string[] inventoryTransactionsColType = { "AUTOINCREMENT", "int", "int", "text", "text", "text", "int", "datetime", "double", "double", "double" };
            //  InvoiceRegister table
            string[] invoiceRegisterColumns = { "P_Id", "SaleDate", "Invoice", "CustomerText", "Employee", "Customer", "JobNotes", "SaleAmount", "Cost", "Net" };
            string[] invoiceRegisterColType = { "AUTOINCREMENT", "datetime", "int", "text", "int", "int", "text", "Currency", "Currency", "Currency" };
            //  SALESLINES table
            string[] salesLinesColumns = { "P_Id", "Invoice", "Qty", "Pc", "Sku", "Description", "Price", "Extend", "ProductNotes" };
            string[] salesLinesColType = { "AUTOINCREMENT", "int NOT NULL", "double", "int NOT NULL", "int NOT NULL", "TEXT(200)", "Currency", "Currency", "TEXT(200)" };
            //  SALES table
            string[] salesColumns = { "P_Id", "SaleDate", "Invoice", "Employee", "Customer", "JobNotes", "SaleAmount", "SalesTax", "Total", "PaymentReceived", "NetAmount" };
            string[] salesColType = { "AUTOINCREMENT", "datetime", "int NOT NULL", "int", "int", "TEXT(200)", "Currency", "Currency", "Currency", "Currency", "Currency" };

            //Create a new database
            bool blnSuccess = CreateDB(@"c:\test.mdb");

            // Create the tables
            CreateTable("Sales", "Invoice", salesColumns, salesColType, pstrDB);
            CreateTable("SalesLines", "P_Id", salesLinesColumns, salesLinesColType, pstrDB);
            CreateTable("Products", "Pc, Sku", productsColumns, productsColType, pstrDB);
            CreateTable("Customers", "CustomerID", customersColumns, customersColType, pstrDB);
            CreateTable("Employees", "EmployeeID", employeeColumns, employeeColType, pstrDB);
            DataSet employees = new DataSet();
            employees.Tables.Add("Employees");
            for (int i = 0; i < employeeColumns.Length; i++)
            {
                Type mysystype = (3).GetType();

                switch (employeeColType[i])
                {
                    case "AUTOINCREMENT":
                        {
                            mysystype = (3).GetType();
                            break;
                        }
                    case "int":
                        {
                            mysystype = (3).GetType();
                            break;
                        }
                    case "text":
                        {
                            mysystype = ("blaa").GetType();
                            break;
                        }
                    case "datetime":
                        {
                            mysystype = new DateTime().GetType();
                            break;
                        }
                }
                employees.Tables[0].Columns.Add(employeeColumns[i], mysystype);
            }

            #region add each employee manually...
            employees.Tables[0].Rows.Add(new object[] { null, 1, "CAROLYN" });
            employees.Tables[0].Rows.Add(new object[] { null, 2, "Ari" });
            employees.Tables[0].Rows.Add(new object[] { null, 3, "CLAY" });
            employees.Tables[0].Rows.Add(new object[] { null, 4, "MATT" });
            employees.Tables[0].Rows.Add(new object[] { null, 5, "Shelly" });
            employees.Tables[0].Rows.Add(new object[] { null, 6, "Heather" });
            employees.Tables[0].Rows.Add(new object[] { null, 7, "Stephanie" });
            employees.Tables[0].Rows.Add(new object[] { null, 8, "Lee" });
            employees.Tables[0].Rows.Add(new object[] { null, 9, "KEVIN" });
            employees.Tables[0].Rows.Add(new object[] { null, 10, "Marie" });
            employees.Tables[0].Rows.Add(new object[] { null, 11, "Bob Manhatton" });
            employees.Tables[0].Rows.Add(new object[] { null, 12, "Rob" });
            employees.Tables[0].Rows.Add(new object[] { null, 13, "Jim" });
            employees.Tables[0].Rows.Add(new object[] { null, 14, "HEATHER/KEVIN" });
            employees.Tables[0].Rows.Add(new object[] { null, 15, "LEAH" });
            employees.Tables[0].Rows.Add(new object[] { null, 16, "JEAN" });
            employees.Tables[0].Rows.Add(new object[] { null, 17, "" });
            employees.Tables[0].Rows.Add(new object[] { null, 18, "PATRICK" });
            employees.Tables[0].Rows.Add(new object[] { null, 19, "Krissy" });  //STEPHEN
            employees.Tables[0].Rows.Add(new object[] { null, 20, "STEPHEN" });  // STEPHEN 
            employees.Tables[0].Rows.Add(new object[] { null, 21, "Bob" });  //  Bob



            UploadToDbsClean(employees, "Employees", @"C:\test.mdb");  // fix me... pstrDB not accesible here...

            #endregion


            CreateTable("InvoiceRegister", "Invoice", invoiceRegisterColumns, invoiceRegisterColType, pstrDB);
            CreateTable("InventoryTransactions", "P_Id", inventoryTransactionsColumns, inventoryTransactionsColType, pstrDB);

            string[] chartofAccountsColumns = { "P_Id", "AccountNumber", "ExtInterface", "AccountTitle", "AccountType", "Division" };
            string[] chartofAccountsColType = { "AUTOINCREMENT", "int", "text", "text", "text", "int" };
            CreateTable("ChartOfAccounts", "AccountNumber", chartofAccountsColumns, chartofAccountsColType, pstrDB);

            // PRODUCT CATEGORY TABLES

            string[] productCategoriesColumns = { "P_Id", "Category", "Description", "UnitOfMeasure", "Taxable", "VendorID", "DecimalPrecision", "Inv", "Inc", "Cos", "Type" };
            string[] productCategoriesColType = { "AUTOINCREMENT", "int", "text", "text", "yesno", "int", "int"                                , "int", "int", "int", "text" };
            CreateTable("ProductCategories", "Category", productCategoriesColumns, productCategoriesColType, pstrDB);

            string[] productCatPromptColumns = { "P_Id", "Category", "PromptNbr", "PromptText" };
            string[] productCatPromptColType = { "AUTOINCREMENT", "int", "int", "text"};
            CreateTable("ProductCatPrompts", "", productCatPromptColumns, productCatPromptColType, pstrDB);

            string[] productCatPriceColumns = { "P_Id", "Category", "PriceLevel", "MarkUp", "RoundingFactor" };
            string[] productCatPriceColType = { "AUTOINCREMENT", "int", "int", "double", "double" };         // test if percentage is correct phrase term
            CreateTable("ProductCatPrices", "", productCatPriceColumns, productCatPriceColType, pstrDB);


            //  VENDOR TABLES

            string[] vendorColumns = { "P_Id", "VendorID", "OriginDate", "Name", "AddressTextToBeObsoleted", "ExternalAccount", "Phone1", "Phone2", "Fax", "LiabilityAccount", "VendorClass" };
            string[] vendorColType = { "AUTOINCREMENT", "int", "DateTime", "text", "text", "text", "text", "text", "text", "int", "int" };
            CreateTable("Vendors", "VendorID", vendorColumns, vendorColType, pstrDB);

            string[] vendorAddresColumns = { "P_Id", "VendorID", "Address1", "Address2", "City", "Zip", "State", "Country", "HiddenPIDLink" };
            string[] vendorAddresColType = { "AUTOINCREMENT", "int", "text", "text", "text", "text", "text", "text", "int" };
            CreateTable("VendorAddresses", "", vendorAddresColumns, vendorAddresColType, pstrDB);

            string[] vendorLedgerTableColumns = { "P_Id", "VendorID", "SeqNbr", "LedgerAccount" };
            string[] vendorLedgerTableColType = { "AUTOINCREMENT", "int", "int", "int" };
            CreateTable("VendorLedgerTable", "", vendorLedgerTableColumns, vendorLedgerTableColType, pstrDB);
            
            string[] vendorPaymentInfoColumns = { "P_Id", "VendorID",  "PaymentPriority", "PaymentTerms", "DiscountTerms" };
            string[] vendorPaymentInfoColType = { "AUTOINCREMENT", "int",  "int", "text", "text" };
            CreateTable("VendorPaymentInfo", "", vendorPaymentInfoColumns, vendorPaymentInfoColType, pstrDB);

            string[] vendorContactColumns = { "P_Id", "VendorID", "SeqNbr", "ContactID", "FirstName", "LastName", "Phone1", "Phone2", "Fax", "Email" };
            string[] vendorContactColType = { "AUTOINCREMENT", "int", "int", "int", "text", "text", "text", "text", "text", "text" };
            CreateTable("VendorContacts", "ContactID", vendorContactColumns, vendorContactColType, pstrDB);

            string[] vendorClassColumns = { "P_Id", "VendorClass", "Description" };
            string[] vendorClassColType = { "AUTOINCREMENT", "int", "text" };
            CreateTable("VendorClasses", "VendorClass", vendorClassColumns, vendorClassColType, pstrDB);


            //  ACCOUNTS PAYABLE TABLES

            //  holds vouchers                //                    trn# OR CheckNbr if type=PMT                        Voucher or Credit Memo
            string[] accountsPayableColumns = { "P_Id", "VendorID", "TransactionNumber", "Invoice", "TransactionDate", "TransactionType", "Amount", "Comment", "CheckNumber" };
            string[] accountsPayableColType = { "AUTOINCREMENT", "int", "int", "int", "datetime", "text", "currency", "text", "int" };
            CreateTable("AccountsPayableLedger", "TransactionNumber", accountsPayableColumns, accountsPayableColType, pstrDB);

            // CHECKBOOK TABLES
            //                                                                                                                              if paid to a vendor... search A/P ledger for this field
            string[] checkbookRegisterColumns = { "P_Id", "CheckNumber", "DatePrinted", "DateCleared", "Description", "Deposits", "CheckNum", "VendorID" };
            string[] checkbookRegisterColType = { "AUTOINCREMENT", "int", "DateTime", "DateTime", "Text", "text", "int", "int" };
            CreateTable("CheckbookRegister", "CheckNumber", checkbookRegisterColumns, checkbookRegisterColType, pstrDB);
        }


        /// <summary>
        /// This function skips the first item in columnNames and colTypes because they are assumed to NOT be parsed.
        /// (That's because they get autocalculated by the database.)
        /// </summary>
        /// <param name="mydataset"></param>
        /// <param name="tableName"></param>
        /// <param name="columnNames"></param>
        /// <param name="colTypes"></param>
        private static void UploadToDbs(DataSet mydataset, string tableName, string[] columnNames, string[] colTypes)
        {
            OleDbConnection myConnection = new OleDbConnection();
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter();
            string pstrDB = @"c:\test.mdb";           // string pointing to database location
            OleDbCommand insertCommand = new OleDbCommand();
            insertCommand.Connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB);



            string insertStr;
            insertStr = "INSERT INTO " + tableName + "( " + columnNames[1];         // note we're skipping [0]

            for (int i = 2; i < columnNames.Length; i++)                            // we're doing all except [0] by the end of it all...
            {
                insertStr += ", " + columnNames[i];
            }
            insertStr += ") VALUES (?";

            for (int i = 2; i < columnNames.Length; i++)                               // we're skipping the first one since it's above (the ?) and skipping one because we forget the first column
            {
                insertStr += ",?";
            }
            insertStr += ")";


            insertCommand.CommandText = insertStr;

            for (int i = 1; i < colTypes.Length; i++)
            {
                switch (colTypes[i].ToLower())
                {
                    case "autoincrement":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Variant, 40, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            //throw new Exception("You (the user) were never supposed to see this error message.  Something has gone slightly wrong.  Call for hlep!");
                            break;
                        }
                    case "datetime":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Date, 30, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "int":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Integer, 30, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "int not null":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Integer, 30, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "currency":
                        {
                            // FIXME  CRAP!  What if this messes things up (see that I put size as one to make it obviouse if it will mess with currency...)
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Currency, 1, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "text":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Variant, 200, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "yesno":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Boolean, 200, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    default:       // assume text...
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Variant, 200, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                }
            }


            myConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB;
            insertCommand.Connection = myConnection;

            myDataAdapter.InsertCommand = insertCommand;

            myConnection.Open();
            try
            {
                myDataAdapter.Update(mydataset, tableName);
            }
            catch (OleDbException exp)
            {
                //if (exp.ErrorCode != -2147467259)         // happens when no insert permission on sales...
                    //MessageBox.Show("Call IT, Unexpected error:  " + exp.Message);

                throw new Exception("Bad news, something went fail: " + exp.Message);
            }
            myConnection.Close();
        }


        /// <summary>
        /// Uploads (inserts) everything in the DATASET under the specified TABLENAME, putting it into the DATABASE
        /// 
        /// This function skips the first item in the DATASET because they are assumed to NOT be parsed "P_Id" autoincrementers...
        /// (you know, because they get autocalculated by the database.)
        /// </summary>
        /// <param name="mydataset"></param>
        /// <param name="tableName"></param>
        /// <param name="columnNames"></param>
        /// <param name="colTypes"></param>
        private static void UploadToDbsClean(DataSet mydataset, string tableName, string pstrDB)
        {
            DataColumnCollection columns = mydataset.Tables[tableName].Columns;

            OleDbConnection myConnection = new OleDbConnection();
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter();
            //string pstrDB = @"c:\test.mdb";           // string pointing to database location
            OleDbCommand insertCommand = new OleDbCommand();
            insertCommand.Connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB);

            //  Parameterized Method

            #region SQL Parameterized statement
            string refstringz = "INSERT INTO " + tableName + "( " + columns[1].ToString();

            for (int l = 2; l < columns.Count; l++)                            // we're doing all except [0] by the end of it all...
            {
                refstringz += ", " + columns[l].ToString();
            }

            refstringz += ") VALUES (";

            refstringz += "?";

            for (int l = 2; l < columns.Count; l++)                               // we're skipping the first one since it's above (the ?) and skipping one because we forget the first column
            {

                refstringz += ", ?";
            }
            refstringz += ")";

            insertCommand.CommandText = refstringz;
            #endregion
            string SQLStatementSetup = refstringz;

            #region Parameter Section

            for (int i = 1; i < columns.Count; i++)  // for each column minus the first
            {
                string colName = columns[i].ColumnName;

                string iDataTypeOfColumn = columns[i].DataType.Name.ToLower();
                switch (iDataTypeOfColumn)
                {
                    case "autoincrement":
                        {
                            //throw new Exception("You (the user) were never supposed to see this error message.  Something has gone slightly wrong.  Call for hlep!");
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 40, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "datetime":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Date, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "int32":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Integer, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "decimal":
                        {
                            // FIXME  CRAP!  What if this messes things up (see that I put size as one to make it obviouse if it will mess with currency...)
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Currency, 1, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "double":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Double, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "string":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "yesno":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Boolean, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "boolean":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Boolean, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    default:       // assume text...
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                }
            }

            #endregion
            OleDbParameterCollection ParametersAreSetup = insertCommand.Parameters;
            

            #region Fill values Loop
            insertCommand.Connection.Open();
            for (int i = 0; i < mydataset.Tables[tableName].Rows.Count; i++)      // for every ROW
            {
                DataRow iRow = mydataset.Tables[tableName].Rows[i];

                // for every COLUMN
                // columns.Count should be greater than Parameters count by exactly one because we're supposed to skip the first column
                for (int l = 1; l < columns.Count; l++)
                {
                    object iField = iRow[l];
                    insertCommand.Parameters[l - 1].Value = iField;
                }
                try
                {
                    insertCommand.ExecuteNonQuery();
                }
                catch (OleDbException exp)
                {
                    //if (exp.ErrorCode != -2147467259)
                        //MessageBox.Show("Some Unexpected Error \n\n" + exp.Message);
                    //MessageBox.Show("Some Error \n\n" + exp.Message);

                    //MessageBox.Show("I don't know how, but workpro consists of duplicate registers... \n\n " + exp.Message, "Duplicate Record Error");
                }//  -2147467259
                
            }
            insertCommand.Connection.Close();
            #endregion
        }


        /// <summary>
        /// Uploads (inserts) everything in the DATASET under the specified TABLENAME, putting it into the DATABASE
        /// 
        /// This function DOES NOT SKIP the first column in the DATASET...
        /// </summary>
        /// <param name="mydataset"></param>
        /// <param name="tableName"></param>
        /// <param name="columnNames"></param>
        /// <param name="colTypes"></param>
        private static void UploadToDbsCleanNoSkip(DataSet mydataset, string tableName, string pstrDB)
        {
            DataColumnCollection columns = mydataset.Tables[tableName].Columns;

            OleDbConnection myConnection = new OleDbConnection();
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter();
            //string pstrDB = @"c:\test.mdb";           // string pointing to database location
            OleDbCommand insertCommand = new OleDbCommand();
            insertCommand.Connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB);

            //  Parameterized Method

            #region SQL Parameterized statement
            string refstringz = "INSERT INTO " + tableName + "( " + columns[0].ToString();

            for (int l = 1; l < columns.Count; l++)                            // we're doing all columns
            {
                refstringz += ", " + columns[l].ToString();
            }

            refstringz += ") VALUES (";
            refstringz += "?";
            for (int l = 1; l < columns.Count; l++)                               // we're skipping the first one since it's above (the ? that didn't proceed a comma)
            {
                refstringz += ", ?";
            }
            refstringz += ")";

            insertCommand.CommandText = refstringz;
            #endregion
            string SQLStatementSetup = refstringz;

            #region Parameter Section

            for (int i = 0; i < columns.Count; i++)  // for each column
            {
                string colName = columns[i].ColumnName;

                string iDataTypeOfColumn = columns[i].DataType.Name.ToLower();
                switch (iDataTypeOfColumn)
                {
                    case "autoincrement":
                        {
                            //MessageBox.Show("You (the user) were never supposed to see this error message.  Something has gone slightly wrong.  Call for hlep!");
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 40, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "datetime":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Date, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "int32":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Integer, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "decimal":
                        {
                            // FIXME  CRAP!  What if this messes things up (see that I put size as one to make it obviouse if it will mess with currency...)
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Currency, 1, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "double":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Double, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "string":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "yesno":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Boolean, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "boolean":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Boolean, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    default:       // assume text...
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                }
            }

            #endregion
            OleDbParameterCollection ParametersAreSetup = insertCommand.Parameters;


            #region Fill values Loop
            insertCommand.Connection.Open();
            for (int i = 0; i < mydataset.Tables[tableName].Rows.Count; i++)      // for every ROW
            {
                DataRow iRow = mydataset.Tables[tableName].Rows[i];

                // for every COLUMN
                for (int l = 0; l < columns.Count; l++)
                {
                    object iField = iRow[l];
                    insertCommand.Parameters[l - 1].Value = iField;
                }
                try
                {
                    insertCommand.ExecuteNonQuery();
                }
                catch (OleDbException exp)
                {
                    //if (exp.ErrorCode != -2147467259)
                    //MessageBox.Show("Some Unexpected Error \n\n" + exp.Message);
                    //MessageBox.Show("Some Error \n\n" + exp.Message);

                    //MessageBox.Show("I don't know how, but workpro consists of duplicate registers... \n\n " + exp.Message, "Duplicate Record Error");
                }//  -2147467259

            }
            insertCommand.Connection.Close();
            #endregion
        }


        /// <summary>
        /// this inverts the qty...  only for use with Inventory transaction journal
        /// </summary>
        /// <param name="rowsToInsert"></param>
        /// <param name="tableName"></param>
        /// <param name="pstrDB"></param>
        private void UploadToDbsCleanFromIRJ(DataRow[] rowsToInsert, string tableName, string pstrDB)
        {
            DataSet convertedRowsToSet = new DataSet();
            //DataTable myDt = new DataTable(tableName);
            // Below makes a new table in the image of tableName.  It's named tableName...
            convertedRowsToSet.Tables.Add(RunQuery("SELECT top 1 * FROM " + tableName, tableName).Tables[tableName].Clone());

            foreach (DataRow dr in rowsToInsert)
            {
                object math = (object)dr["Qty"];// *-1;
                double qtyDec = Convert.ToDouble(math.ToString());
                qtyDec = qtyDec * -1;
                dr["Qty"] = (object)qtyDec;

                convertedRowsToSet.Tables[tableName].ImportRow(dr); //.Rows.Add(dr);
            }
            UploadToDbsClean(convertedRowsToSet, tableName, pstrDB);
        }



        /// <summary>
        /// Uploads everything in the DATASET under the specified TABLENAME, putting it into the DATABASE
        /// 
        /// this does the whole dataset at once.
        /// 
        /// This function skips the first item in the DATASET because they are assumed to NOT be parsed "P_Id" autoincrementers...
        /// (you know, because they get autocalculated by the database.)
        /// </summary>
        /// <param name="mydataset"></param>
        /// <param name="tableName"></param>
        /// <param name="columnNames"></param>
        /// <param name="colTypes"></param>
        private static void UploadToDbsBatch(DataSet mydataset, string tableName, string pstrDB)
        {
            DataColumnCollection columns = mydataset.Tables[tableName].Columns;

            OleDbConnection myConnection = new OleDbConnection();
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter();
            //string pstrDB = @"c:\test.mdb";           // string pointing to database location
            OleDbCommand insertCommand = new OleDbCommand();
            insertCommand.Connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB);

            //  Parameterized Method

            #region SQL Parameterized statement
            string refstringz = "INSERT INTO " + tableName + "( " + columns[1].ToString();

            for (int l = 2; l < columns.Count; l++)                            // we're doing all except [0] by the end of it all...
            {
                refstringz += ", " + columns[l].ToString();
            }

            refstringz += ") VALUES (";

            refstringz += "?";

            for (int l = 2; l < columns.Count; l++)                               // we're skipping the first one since it's above (the ?) and skipping one because we forget the first column
            {

                refstringz += ", ?";
            }
            refstringz += ")";

            insertCommand.CommandText = refstringz;
            #endregion
            string SQLStatementSetup = refstringz;

            #region Parameter Section

            for (int i = 1; i < columns.Count; i++)  // for each column minus the first
            {
                string colName = columns[i].ColumnName;

                string iDataTypeOfColumn = columns[i].DataType.Name.ToLower();
                switch (iDataTypeOfColumn)
                {
                    case "autoincrement":
                        {
                            throw new Exception("You (the user) were never supposed to see this error message.  Something has gone slightly wrong.  Call for hlep!");
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 40, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "datetime":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Date, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "int32":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Integer, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "decimal":
                        {
                            // FIXME  CRAP!  What if this messes things up (see that I put size as one to make it obviouse if it will mess with currency...)
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Currency, 1, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "double":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Double, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "string":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "yesno":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Boolean, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "boolean":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Boolean, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    default:       // assume text...
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                }
            }

            #endregion
            OleDbParameterCollection ParametersAreSetup = insertCommand.Parameters;
            myDataAdapter.InsertCommand = insertCommand;

            
            insertCommand.Connection.Open();
            myDataAdapter.Update(mydataset, tableName);
            insertCommand.Connection.Close();
        }


        /// <summary>
        /// Use this one if all you need to upload to the .mdb file are integer values.  
        /// Store the integers in arrays.  Store the table names in arrays (with an order corresponding to the integer arrays)
        /// </summary>
        /// <param name="mydataset"></param>
        /// <param name="tableName"></param>
        /// <param name="columnNames"></param>
        /// <param name="colTypes"></param>
        private static void UploadToDbsDecimals(DataSet mydataset, string tableName, string[] columnNames, decimal[] dataValues)
        {
            string dataType = "Currency";

            OleDbConnection myConnection = new OleDbConnection();
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter();

            OleDbCommand insertCommand = new OleDbCommand();
            //        SET UP DATABASE CONNECTION JAZZ

            string pstrDB = @"c:\test.mdb";           // string pointing to database location


            string insertStr;
            insertStr = "INSERT INTO " + tableName + "( " + columnNames[1];

            for (int i = 2; i < columnNames.Length; i++)
            {
                insertStr += ", " + columnNames[i];
            }
            insertStr += ") VALUES (?";

            for (int i = 2; i < columnNames.Length; i++)
            {
                insertStr += ",?";
            }
            insertStr += ")";


            insertCommand.CommandText = insertStr;

            for (int i = 1; i < columnNames.Length; i++)
            {
                switch (dataType.ToLower())
                {
                    case "autoincrement":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Variant, 40, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "datetime":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Date, 30, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName

                            break;
                        }
                    case "int":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Integer, 30, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName

                            break;
                        }
                    case "int not null":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Integer, 30, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName

                            break;
                        }
                    case "currency":
                        {
                            // FIXME  CRAP!  What if this messes things up (see that I put size as one to make it obviouse if it will mess with currency...)
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Currency, 1, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName

                            break;
                        }
                    default:       // assume text...
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(columnNames[i], OleDbType.Variant, 200, columnNames[i]));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                }
            }


            myConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB;
            insertCommand.Connection = myConnection;

            myDataAdapter.InsertCommand = insertCommand;

            myConnection.Open();

            myDataAdapter.Update(mydataset, tableName);

            myConnection.Close();
        }

        

        public static bool CreateDB(string pstrDB)
        {
            try
            {
                Catalog cat = new Catalog();
                string strCreateDB = "";

                strCreateDB += "Provider=Microsoft.Jet.OLEDB.4.0;";
                strCreateDB += "Data Source=" + pstrDB + ";";
                strCreateDB += "Jet OLEDB:Engine Type=5";
                cat.Create(strCreateDB);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        internal static void CreateTable(string TableName, string PrimaryKey, string[] DailySalesColumns, string[] DailySalesColType, string pstrDB)
        {
            #region Paramiter validity checks
            bool fail = false;
            // check if daily sales columns and col types are of equal length
            if (DailySalesColumns.Length != DailySalesColType.Length) fail = true;

            if (fail == true)
            {
                throw new Exception("Error in Parameters sent to CreateTable(" + TableName + "...)");
                throw new Exception("Paramiters for CreateTable(...) were inapropriate.");
            }
                    
            #endregion


            //  This creates a brand new table called SampleTable in test.mdb
            string tableName = TableName;
            string[] dailySalesColumns = DailySalesColumns;
            string[] dailySalesColType = DailySalesColType;
            string primaryKey = PrimaryKey;

            if (primaryKey != "")
                primaryKey = ", CONSTRAINT " + tableName + "_PK PRIMARY KEY (" + primaryKey + ")";

            String svlQuery = "CREATE TABLE " + tableName + " ( " + dailySalesColumns[0] + " " + dailySalesColType[0];

            for (int i = 1; i < dailySalesColumns.Length; i++)
            {
                svlQuery += ", " + dailySalesColumns[i] + " " + dailySalesColType[i];
            }
            if (primaryKey != "")
                svlQuery += primaryKey;

            svlQuery += ")";

            // Setup Connection
            OleDbConnection oledbConnection1 = new OleDbConnection();
            oledbConnection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB + ";";
            oledbConnection1.Open();

            // Setup Command
            OleDbCommand oledbCommand1 = new OleDbCommand(svlQuery);
            oledbCommand1.Connection = oledbConnection1;

            //Run Command
            try
            {
                oledbCommand1.ExecuteNonQuery();
            }
            catch (Exception exp)
            {
                throw new Exception("Table " + tableName + " not created because... \n\r" + exp.ToString());
            }
            oledbConnection1.Close();                       // Close connection
        }

        /// <summary>
        /// What ever is in the DATABASE gets put into the TABLE of the DATASET
        /// Loads up mydataset with tableName from Database
        /// </summary>
        /// <param name="pstrDB">DATABASE</param>
        /// <param name="tableName">TABLE</param>
        /// <param name="mydataset">DATASET</param>
        public DataSet FillDataSet(string tableName, string pstrDB)
        {
            DataSet returnset = new DataSet();
            String svlQuery = "select * from " + tableName;

            OleDbConnection oledbConnection1 = new OleDbConnection();
            oledbConnection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB + ";";
            OleDbDataAdapter oledbDataAdapter1 = new OleDbDataAdapter(svlQuery, oledbConnection1);

            oledbConnection1.Open();
            oledbDataAdapter1.Fill(returnset, tableName);
            oledbConnection1.Close();
            return returnset;
        }


        



        /// <summary>
        /// Take data from one COLUMN of a DATATABLE and put it in a DATASET
        /// What ever is in the DATABASE under the specified table name gets put into the TABLE of the new DATASET returned 
        /// (under the same table name)
        /// </summary>
        /// <param name="tblName"></param>
        /// <returns></returns>
        private DataSet FillNewDataSet(string tblName, string colName, string pstrDB)
        {
            DataSet returnDataset = new DataSet();
            String svlQuery = "select " + colName + " from " + tblName;           // FIXME  I changed where the SaleDate to an astrik

            OleDbConnection oledbConnection1 = new OleDbConnection();
            oledbConnection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB + ";";
            OleDbDataAdapter oledbDataAdapter1 = new OleDbDataAdapter(svlQuery, oledbConnection1);

            oledbConnection1.Open();
            oledbDataAdapter1.Fill(returnDataset, tblName);
            oledbConnection1.Close();

            return returnDataset;
        }

        /// <summary>
        /// I love this function
        /// It uses the DATASET you pass, checking the TABLE's COLUMN to see if the DATABASE has a matching row, then deletes it
        /// </summary>
        /// <param name="dataSet"></param>
        /// <param name="tableName"></param>
        /// <param name="IDCol"></param>
        /// <param name="pstrDB"></param>
        private void DeleteOverlaps(DataSet dataSet, string tableName, string IDCol, string pstrDB)
        {
            // delete record where exists in dataSet
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            string column = IDCol;
            string refstringz = "DELETE * FROM " + tableName + " WHERE (" + column + " = ";
            OleDbCommand deleteCommand = new OleDbCommand();
            deleteCommand.Connection = new OleDbConnection(MakeConnectionString(pstrDB));

            deleteCommand.Connection.Open();
            foreach (DataRow dr in dataSet.Tables[tableName].Rows)         // for everything in the dataSet..., Delete dr...
            {
                string refstringzCho = refstringz  + dr[IDCol].ToString() + ")";
                deleteCommand.CommandText = refstringzCho;

                deleteCommand.ExecuteNonQuery();
            }
            deleteCommand.Connection.Close();

            string tim = stopwatch.Elapsed.ToString();
            stopwatch.Stop();
        }



        private void deleteTable(string tblNamez)
        {
            String svlQuery = "DELETE FROM " + tblNamez;

            System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(svlQuery,
                new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB + ";"));

            //Run Command
            try
            {
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();
            }
            catch (OleDbException exp)
            {
                throw new Exception("Table not deleted because... \n\r " + exp.Message);
            }
            finally
            {
                cmd.Connection.Close();
            }
        }




        /// <summary>
        /// This command runs a query...  
        /// Pass a select query to this function (along with a tableName) and the data selected will be put in the tableName specified.
        /// </summary>
        /// <param name="pstrDB">DATABASE</param>
        /// <param name="tableName">TABLE</param>
        /// <param name="mydataset">DATASET</param>
        public DataSet RunQuery(string svl, string returnTableName)
        {
            String svlQuery = svl;
            DataSet returnSet = new DataSet();

            OleDbConnection oledbConnection1 = new OleDbConnection();
            oledbConnection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB + ";";
            OleDbDataAdapter oledbDataAdapter1 = new OleDbDataAdapter(svlQuery, oledbConnection1);

            oledbConnection1.Open();
            oledbDataAdapter1.Fill(returnSet, returnTableName);
            oledbConnection1.Close();

            return returnSet;
        }


        public void RunNonQuerry(string svlQuery)
        {
            System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand(svlQuery,
                new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB + ";"));

            cmd.Connection.Open();
            cmd.ExecuteNonQuery();
            cmd.Connection.Close();
        }

        /// <summary>
        /// Search for the first occurance of SEARCHPATTERN within the TABLE, under the specified COLUMN
        /// and this will return that whole row.
        /// 
        /// This is used for searching the Employee table for matching the Employee text of the invoice register
        /// with the employee number
        /// </summary>
        /// <param name="SearchPattern"></param>
        /// <param name="TableName"></param>
        /// <param name="ColumnToSearch"></param>
        /// <returns>the matching DataRow.  If record not found, return null.</returns>
        private DataRow SearchDatabaseForFirst(string SearchPattern, string TableName, string ColumnToSearch)
        {
            DataSet myEmployees = FillNewDataSet(TableName, "*", pstrDB);
            DataTable employeeTable = myEmployees.Tables[TableName];

            for (int i = 0; i < employeeTable.Rows.Count; i++)
            {
                if (employeeTable.Rows[i][ColumnToSearch].ToString().ToLower() == SearchPattern.ToLower())
                {
                    return employeeTable.Rows[i];
                }
            }

            return null;
        }

        /// <summary>
        /// Search for the first occurance of SEARCHPATTERN within the TABLE, under the specified COLUMN
        /// and this will return that whole row.
        /// 
        /// This is used for searching the Employee table for matching the Employee text of the invoice register
        /// with the employee number
        /// </summary>
        /// <param name="SearchPattern"></param>
        /// <param name="TableName"></param>
        /// <param name="ColumnToSearch"></param>
        /// <returns>the matching DataRow.  If record not found, return null.</returns>
        private bool recordExists(string SearchPattern, DataSet dt, string ColumnToSearch)
        {
            DataTable myTable = dt.Tables[0];

            for (int i = 0; i < myTable.Rows.Count; i++)
            {
                if (myTable.Rows[i][ColumnToSearch].ToString().ToLower() == SearchPattern.ToLower())
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Search for the first occurance of SEARCHPATTERN within the TABLE, under the specified COLUMN
        /// and this will return that whole row.
        /// 
        /// This is used for searching the Employee table for matching the Employee text of the invoice register
        /// with the employee number
        /// </summary>
        /// <param name="SearchPattern"></param>
        /// <param name="TableName"></param>
        /// <param name="ColumnToSearch"></param>
        /// <returns>the matching DataRow.  If record not found, return null.</returns>
        private bool itemExists(int SearchPattern, int[] intList)
        {
            for (int i = 0; i < intList.Length; i++)
            {
                if (intList[i] == SearchPattern)
                {
                    return true;
                }
            }
            return false;
        }


        private DataTable SearchDatabaseForRows(string SearchPattern, string TableName, string ColumnToSearch, string pstrDB)
        {
            DataSet wasInDataBase = FillNewDataSet(TableName, "*", pstrDB);
            DataTable wasInTable = wasInDataBase.Tables[TableName];
            DataTable hits = wasInTable.Clone();

            for (int i = 0; i < wasInTable.Rows.Count; i++)
            {
                if (wasInTable.Rows[i][ColumnToSearch].ToString().ToLower() == SearchPattern.ToLower())
                {
                    hits.Rows.Add(wasInTable.Rows[i]);
                }
            }

            if (hits.Rows.Count > 0)
                return hits;
            else
                return null;
        }

        /// <summary>
        /// Updates all Columns matching the ID COLUMN with the values in the DatasetWithFieldsToUpdateFrom...
        /// 
        /// This function only works for the invoice register... didn't want to waste time properly encapsulating/rigging it
        /// 
        /// The first column is assumed to be the IDENTIFIEING column
        /// The second column is a datetime field which will be updated
        /// </summary>
        /// <param name="DatasetWithFieldsToUpdateFrom"></param>
        /// <param name="TableName"></param>
        /// <param name="ColumnNames"></param>
        /// <param name="databasePath"></param>
        private void UpdateDatabaseFieldsDated(DataSet DatasetWithFieldsToUpdateFrom, string TableName, string[] ColumnNames, string databasePath)
        {
            OleDbCommand updateCommand = new OleDbCommand();
            updateCommand.Connection = new OleDbConnection(MakeConnectionString(databasePath));
            OleDbDataAdapter oledbDataAdapter1 = new OleDbDataAdapter();

            updateCommand.Connection.Open();


            for (int i = 0; i < DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows.Count; i++)
            {
                string refstringz = " UPDATE " + TableName + " SET " + ColumnNames[1] + " = #" + DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows[i][ColumnNames[1]] + "#";
                refstringz += " WHERE (" + ColumnNames[0] + " = " + DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows[i][ColumnNames[0]] + ")";
                updateCommand.CommandText = refstringz;

                updateCommand.ExecuteNonQuery();
            }
            updateCommand.Connection.Close();
        }


        /// <summary>
        /// First column in Column names is the ID Column and it's data type seperated by comma...
        /// Following columns are updated.  They each must be followed in comma by as datatype such as "Column1, int".  if there is no comma, then int is assumed...
        /// 
        /// must have atleast two columns in ColumnNames.  
        /// </summary>
        /// <param name="DatasetWithFieldsToUpdateFrom"></param>
        /// <param name="TableName"></param>
        /// <param name="ColumnNames"></param>
        /// <param name="databasePath"></param>
        private void UpdateDatabaseFields(DataSet DatasetWithFieldsToUpdateFrom, string TableName, string[] ColumnNames, string databasePath)
        {
            OleDbCommand updateCommand = new OleDbCommand();
            updateCommand.Connection = new OleDbConnection(MakeConnectionString(databasePath));
            OleDbDataAdapter oledbDataAdapter1 = new OleDbDataAdapter();

            //  "Name, int"
            #region strip dataTypes
            string[] DataTypes = new string[ColumnNames.Length];

            for (int i = 0; i < ColumnNames.Length; i++)
            {
                try
                {
                    if (ColumnNames[i].IndexOf(',') != -1)
                        DataTypes[i] = ColumnNames[i].Substring(ColumnNames[i].IndexOf(',') + 1).Trim();
                    else
                        DataTypes[i] = "int";
                    ColumnNames[i] = ColumnNames[i].Substring(0, ColumnNames[i].IndexOf(','));
                }
                catch (Exception)
                {
                    throw new Exception("Error stripping data types from ColumnNames which was passed to UpdateDatabaseFields.  Plz check your code and make sure you didn't mess up and forget a column type.");
                }
            }
            #endregion

            updateCommand.Connection.Open();

            string TOKEN = "#";

            //string refstringz = " UPDATE " + TableName + " SET ";


            for (int i = 0; i < DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows.Count; i++)  // for each row in dataset
            {
                #region setup special token for datetime for first column...
                if (DataTypes[1].ToLower() == "datetime")
                    TOKEN = "#";
                else
                    TOKEN = "";
                #endregion

                string refstringz = " UPDATE " + TableName + " SET " + ColumnNames[1] + " = " + TOKEN + DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows[i][ColumnNames[1]] + TOKEN + " ";
                // for each column to be updated
                if (ColumnNames.Length > 2)
                    for (int l = 2; l < ColumnNames.Length; l++)
                    {
                        #region setup special token for datetime
                    if (DataTypes[l].ToLower() == "datetime")
                        TOKEN = "#";
                    else
                        TOKEN = "";
                    #endregion

                        refstringz += ", " + ColumnNames[l] + " = " + TOKEN + DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows[i][ColumnNames[l]] + TOKEN + " ";
                    }

                refstringz += " WHERE (" + ColumnNames[0] + " = " + DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows[i][ColumnNames[0]] + ")";
                updateCommand.CommandText = refstringz;

                updateCommand.ExecuteNonQuery();
            }
            updateCommand.Connection.Close();
        }

        /// <summary>
        /// First column in Column names is the ID Column and it's data type seperated by comma...
        /// Following columns are updated.  They each must be followed in comma by as datatype such as "Column1, int".  if there is no comma, then int is assumed...
        /// 
        /// must have atleast two columns in ColumnNames.  
        /// </summary>
        /// <param name="DatasetWithFieldsToUpdateFrom"></param>
        /// <param name="TableName"></param>
        /// <param name="UpdateColumns"></param>
        /// <param name="databasePath"></param>
        private void UpdateDatabaseFields(DataSet DatasetWithFieldsToUpdateFrom, string TableName, string[] IDColumns, string[] UpdateColumns, string databasePath)
        {
            OleDbCommand updateCommand = new OleDbCommand();
            updateCommand.Connection = new OleDbConnection(MakeConnectionString(databasePath));
            OleDbDataAdapter oledbDataAdapter1 = new OleDbDataAdapter();

            //  "Name, int"
            #region strip dataTypes
            string[] DataTypes = new string[UpdateColumns.Length];

            for (int i = 0; i < UpdateColumns.Length; i++)
            {
                try
                {
                    if (UpdateColumns[i].IndexOf(',') != -1)
                    {
                        DataTypes[i] = UpdateColumns[i].Substring(UpdateColumns[i].IndexOf(',') + 1).Trim();
                        UpdateColumns[i] = UpdateColumns[i].Substring(0, UpdateColumns[i].IndexOf(','));
                    }
                    else
                        DataTypes[i] = "int";
                }
                catch (Exception)
                {
                    throw new Exception("Error stripping data types from ColumnNames which was passed to UpdateDatabaseFields.  Plz check your code and make sure you didn't mess up and forget a column type.");
                }
            }
            #endregion

            #region strip IDColumn dataTypes
            string[] IDDataTypes = new string[IDColumns.Length];

            for (int i = 0; i < IDColumns.Length; i++)
            {
                try
                {
                    if (IDColumns[i].IndexOf(',') != -1)
                    {
                        IDDataTypes[i] = IDColumns[i].Substring(IDColumns[i].IndexOf(',') + 1).Trim();
                        IDColumns[i] = IDColumns[i].Substring(0, IDColumns[i].IndexOf(','));
                    }
                    else
                        IDDataTypes[i] = "int";
                }
                catch (Exception)
                {
                    throw new Exception("Error stripping data types from ColumnNames which was passed to UpdateDatabaseFields.  Plz check your code and make sure you didn't mess up and forget a column type.");
                }
            }
            #endregion

            updateCommand.Connection.Open();

            string TOKEN = "#";


            for (int i = 0; i < DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows.Count; i++)  // for each row in dataset
            {
                #region setup special token for datetime for first column...
                if (DataTypes[0].ToLower() == "datetime")
                    TOKEN = "#";
                else if (DataTypes[0].ToLower() == "text")
                    TOKEN = "'";
                else
                    TOKEN = "";
                #endregion

                string refstringz = " UPDATE " + TableName + " SET " + UpdateColumns[0] + " = " + TOKEN + DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows[i][UpdateColumns[0]] + TOKEN + " ";
                // for each column to be updated
                if (UpdateColumns.Length > 1)
                    for (int l = 1; l < UpdateColumns.Length; l++)
                    {
                        #region setup special token for datetime
                        if (DataTypes[l].ToLower() == "datetime")
                            TOKEN = "#";
                        else if (DataTypes[l].ToLower() == "text")
                            TOKEN = "'";
                        else
                            TOKEN = "";
                        #endregion

                        refstringz += ", " + UpdateColumns[l] + " = " + TOKEN + DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows[i][UpdateColumns[l]] + TOKEN + " ";
                    }

                refstringz += " WHERE (" + IDColumns[0] + " = " + DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows[i][UpdateColumns[0]];

                if (IDColumns.Length > 1)
                    for (int l = 1; l < IDColumns.Length; l++)
                    {
                        #region setup special token for datetime
                        if (IDDataTypes[l].ToLower() == "datetime")
                            TOKEN = "#";
                        else if (IDDataTypes[l].ToLower() == "text")
                            TOKEN = "'";
                        else
                            TOKEN = "";
                        #endregion

                        refstringz += " AND " + IDColumns[l] + " = " + TOKEN + DatasetWithFieldsToUpdateFrom.Tables[TableName].Rows[i][IDColumns[l]] + TOKEN;
                    }

                refstringz += ")";
                updateCommand.CommandText = refstringz;

                updateCommand.ExecuteNonQuery();
            }
            updateCommand.Connection.Close();
        }




        /// <summary>
        /// Uploads (inserts) everything in the DATASET under the specified TABLENAME, putting it into the DATABASE
        /// 
        /// This function skips the first item in the DATASET because they are assumed to NOT be parsed "P_Id" autoincrementers...
        /// (you know, because they get autocalculated by the database.)
        /// </summary>
        /// <param name="mydataset"></param>
        /// <param name="tableName"></param>
        /// <param name="columnNames"></param>
        /// <param name="colTypes"></param>
        private static void UpdateDbsFieldsClean(DataSet mydataset, string tableName, string pstrDB)
        {
            DataColumnCollection columns = mydataset.Tables[tableName].Columns;

            OleDbConnection myConnection = new OleDbConnection();
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter();
            OleDbCommand insertCommand = new OleDbCommand();
            insertCommand.Connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB);


            OleDbCommand updateCommand = new OleDbCommand();
            updateCommand.Connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + pstrDB);

            //  Parameterized Method
                                                        // SET(4)
            updateCommand.CommandText = "UPDATE BookDb SET author = ?, EditionNumber = ?, ISBN = ?, Title = ? " +
                "WHERE (MyCol = ?) " +
                "AND (EditionNumber = ? OR ? IS NULL AND EditionNumber IS NULL) " +
                "AND (Title = ? OR ? IS NULL AND Title IS NULL) " +
                "AND (author = ? OR ? IS NULL AND author IS NULL)";

            updateCommand.CommandText = "UPDATE SalesLines SET Price = ?, Extend = ? " +
                "WHERE (Invoice = ?) " +
                "AND (Pc = ?) " +
                "AND (Sku = ?) " +
                "AND (Qty = ?)" +
                "AND (Price IS NULL)";

            updateCommand.Parameters.Add(new OleDbParameter("Invoice", OleDbType.Integer, 50, "Invoice"));
            updateCommand.Parameters.Add(new OleDbParameter("Pc", OleDbType.Integer, 50, "Pc"));


            #region SQL Parameterized statement
            string refstringz = "UPDATE " + tableName + "SET " + columns[1].ToString();

            for (int l = 2; l < columns.Count; l++)                            // we're doing all except [0] by the end of it all...
            {
                refstringz += ", " + columns[l].ToString();
            }

            refstringz += ") VALUES (";

            refstringz += "?";

            for (int l = 2; l < columns.Count; l++)                               // we're skipping the first one since it's above (the ?) and skipping one because we forget the first column
            {

                refstringz += ", ?";
            }
            refstringz += ")";

            insertCommand.CommandText = refstringz;
            #endregion
            string SQLStatementSetup = refstringz;

            #region Parameter Section

            for (int i = 1; i < columns.Count; i++)  // for each column minus the first
            {
                string colName = columns[i].ColumnName;

                string iDataTypeOfColumn = columns[i].DataType.Name.ToLower();
                switch (iDataTypeOfColumn)
                {
                    case "autoincrement":
                        {
                            throw new Exception("You (the user) were never supposed to see this error message.  Something has gone slightly wrong.  Call for hlep!");
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 40, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "datetime":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Date, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "int32":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Integer, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "decimal":
                        {
                            // FIXME  CRAP!  What if this messes things up (see that I put size as one to make it obviouse if it will mess with currency...)
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Currency, 1, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "double":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Double, 30, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "string":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "yesno":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Boolean, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    case "boolean":
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Boolean, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                    default:       // assume text...
                        {
                            insertCommand.Parameters.Add(new OleDbParameter(colName, OleDbType.Variant, 200, colName));      // UID, datatype, maxLength, DataSet.ColumnName
                            break;
                        }
                }
            }

            #endregion
            OleDbParameterCollection ParametersAreSetup = insertCommand.Parameters;


            #region Fill values Loop
            insertCommand.Connection.Open();
            for (int i = 0; i < mydataset.Tables[tableName].Rows.Count; i++)      // for every ROW
            {
                DataRow iRow = mydataset.Tables[tableName].Rows[i];

                // for every COLUMN
                // columns.Count should be greater than Parameters count by exactly one because we're supposed to skip the first column
                for (int l = 1; l < columns.Count; l++)
                {
                    object iField = iRow[l];
                    insertCommand.Parameters[l - 1].Value = iField;
                }
                try
                {
                    insertCommand.ExecuteNonQuery();
                }
                catch (OleDbException exp)
                {
                    if (exp.ErrorCode != -2147467259)
                        throw new Exception("Some Unexpected Error \n\n" + exp.Message);
                    //MessageBox.Show("Some Error \n\n" + exp.Message);

                    //MessageBox.Show("I don't know how, but workpro consists of duplicate registers... \n\n " + exp.Message, "Duplicate Record Error");
                }//  -2147467259

            }
            insertCommand.Connection.Close();
            #endregion
        }




        /// <summary>
        /// Takes a column and determines if it's a string column or an int or a DateTime
        /// </summary>
        /// <param name="dataColumn"></param>
        /// <returns></returns>
        private OleDbType GetOleDbType(DataColumn dataColumn)
        {
            if (dataColumn.DataType == System.Type.GetType("System.String"))
                return OleDbType.VarWChar;
            if (dataColumn.DataType == System.Type.GetType("System.Int32"))
                return OleDbType.Integer;
            if (dataColumn.DataType == System.Type.GetType("System.DateTime"))
                return OleDbType.Date;

            return OleDbType.VarWChar;
        }

        private string MakeConnectionString(string dbPath)
        {
            return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = " + dbPath;
        }


        /// <summary>
        /// This function deletes all records from the database in specified TABLE where COLUMN matches PATTERN
        /// </summary>
        /// <param name="column"></param>
        /// <param name="p_2"></param>
        /// <param name="pstrDB"></param>
        private void DeleteRecord(string tableName, string column, string pattern, string pstrDB)
        {
            OleDbCommand deleteCommand = new OleDbCommand();
            deleteCommand.Connection = new OleDbConnection(MakeConnectionString(pstrDB));
            OleDbDataAdapter oledbDataAdapter1 = new OleDbDataAdapter();
            //  try OleDbType.Integer if numeric doesn't fly

            string refstringz = "DELETE * FROM " + tableName + " WHERE (";
            refstringz += column + " = " + pattern + ")";
            deleteCommand.CommandText = refstringz;


            deleteCommand.Connection.Open();
            deleteCommand.ExecuteNonQuery();
            deleteCommand.Connection.Close();
        }


        /// <summary>
        /// 1)  ReadDateSpan of RPT to get dateRange[2]
        /// 2)  GetDateRange from in the database..., returning a DataSet with all the dates between that section.
        /// 3)  Count the number of dates which exist in the database and are within the date span...
        /// 4)  run the Delete Between query against the database
        /// 5)  Retun the number of days detected... and deleted... may be multiple records per day...
        /// </summary>
        /// <param name="fil">report file to get DateSpan from</param>
        /// <param name="tableName">Name of table to get overlaps regaurding</param>
        /// <param name="colName">The date column... SaleDate</param>
        /// <returns></returns>
        private int DeleteOverlappingDatesCount(string fil, string tableName, string colName, string pstrDB)
        {
            // check date of file
            DateTime[] dateRange = ReadDateSpan(fil,
                CheckReportType(new FileInfo(fil)));

            List<DateTime> reportRange = GetDateRange(dateRange[0], dateRange[1]);     // retuns a list composed of ALL days within the range

            // read the records
            DataSet tempSet = FillNewDataSet(tableName, colName, pstrDB);  // tblName, ColName
            List<DateTime> theDatabaseDates = new List<DateTime>();            //  Lists all dates already in the database

            for (int i = 0; i < tempSet.Tables[0].Rows.Count; i++)
            {
                theDatabaseDates.Add((DateTime)tempSet.Tables[0].Rows[i][0]);        // put all the dates of the database into a list
            }

            int numberOfOverlaps = 0;

            foreach (DateTime day in reportRange)       // check every day that could exist in the report range
            {
                if (theDatabaseDates.Contains(day))       // to see if that day is already in the database
                {
                    numberOfOverlaps++;
                    // date overlap detected
                }
            }
            string queryStr = "DELETE * FROM " + tableName + " WHERE " + colName + " between #" + dateRange[0].ToShortDateString() + "# AND #" + dateRange[1].ToShortDateString() + "#";
            RunNonQuerry(queryStr);

            return numberOfOverlaps;
        }

        private List<DateTime> GetDateRange(DateTime StartingDate, DateTime EndingDate)
        {
            if (StartingDate > EndingDate)
            {
                return null;
            }
            List<DateTime> rv = new List<DateTime>();
            DateTime tmpDate = StartingDate;
            do
            {
                rv.Add(tmpDate);
                tmpDate = tmpDate.AddDays(1);
            } while (tmpDate <= EndingDate);
            return rv;
        }


        /// <summary>
        /// (For use with inventoryTransactions -> SalesLines ONLY!)
        /// Prepares a DATASET out of the DataRows[] provided it, in the image of the tableName specified of the DATABASE
        /// 
        /// IMPORTANT:  It also changes the polarity of the qty... this is because the qty specified in the inventory transaction journal is oppisite to the qty specified on the sales thing...
        /// </summary>
        /// <param name="rowsToInsert"></param>
        /// <param name="tableName"></param>
        /// <param name="pstrDB"></param>
        /// <returns></returns>
        private DataSet convertRowsToDataSet(DataRow[] rowsToInsert, string tableName, string pstrDB)
        {
            DataSet convertedRowsToSet = new DataSet();
            //DataTable myDt = new DataTable(tableName);
            // Below makes a new table in the image of tableName.  It's named tableName...
            convertedRowsToSet.Tables.Add(RunQuery("SELECT top 1 * FROM " + tableName, tableName).Tables[tableName].Clone());

            foreach (DataRow dr in rowsToInsert)
            {
                object math = (object)dr["Qty"];// *-1;
                double qtyDec = Convert.ToDouble(math.ToString());
                qtyDec = qtyDec * -1;
                dr["Qty"] = (object)qtyDec;

                convertedRowsToSet.Tables[tableName].ImportRow(dr); //.Rows.Add(dr);
            }
            return convertedRowsToSet;
        }




        public reportTypes CheckReportType(FileInfo p)
        {
            string ReportTitle = "";

            TextReader tr2 = new StreamReader(p.FullName);        //Read the file
            char[] FirstSection2 = new char[300];
            tr2.Read(FirstSection2, 0, 300);                 // read just enough of the file to be able to identify the report type
            string firstSection2 = new string(FirstSection2);
            tr2.Close();                                  // close the stream

            #region Style zero
            StreamReader mt = new StreamReader(p.FullName);
            string currentLine;

            mt.ReadLine();

            currentLine = mt.ReadLine();
            mt.ReadLine();
            string atLine = mt.ReadLine();
            string emptLine1 = mt.ReadLine();
            string emptLine2 = mt.ReadLine();
            mt.Close();

            // Get timeLine aka report title line
            string timeLine = "";
            mt = new StreamReader(p.FullName);
            for (int i = 0; i < 5; i++)
            {
                timeLine = mt.ReadLine();
                if (timeLine != null && timeLine.Length > 4)
                    if (timeLine.Substring(0, 4).ToLower() == "time")
                        break;
                timeLine = "";
            }
            mt.Close();

            #endregion


            int pointStart = firstSection2.IndexOf("Time:") + 14;   //  We just jumped past Time: xx:yy:zz
            int pointEnd = firstSection2.IndexOf("\r\n", pointStart);

            if (pointEnd != -1)
            {
                //ReportTitle = firstSection2.Substring(pointStart, pointEnd - pointStart).Trim();
                if (timeLine.Length > 14)
                    ReportTitle = timeLine.Substring(14).Trim();
            }
            else
                return reportTypes.Unknown;

            if (atLine.Length > 22)
                if (atLine.Substring(20, 2) == "at" && emptLine1 == "" && emptLine2 == "")
                    ReportTitle = currentLine.Trim();

            switch (ReportTitle)         //  this switch can't detect the daily reports unfortunately
            {
                case "Chart of Account Master List":
                    {
                        return reportTypes.MasterChartOfAccountsList;
                    }
                case "Accounts Payable Ledger":
                    {
                        return reportTypes.AccountsPayableLedger;
                    }
                case "Checkbook Register":
                    {
                        return reportTypes.CheckBookRegister;
                    }
                case "Vendor Master List":
                    {
                        return reportTypes.MasterVendorList;
                    }
                case "Product Category Master List":
                    {
                        return reportTypes.MasterProductCategoryList;
                    }
                case "Customer Master List":
                    {
                        return reportTypes.MasterCustomerList;
                    }
                case "Inventory Transaction Journal":
                    {
                        return reportTypes.InventoryTransactionJournal;
                    }
                case "Invoice Register by Employee":
                    {
                        return reportTypes.InvoiceRegisterByEmployee;
                    }
                case "Invoice Register by Day":
                    {
                        return reportTypes.InvoiceRegisterByDay;
                    }

                case "CASH RECEIPTS JOURNAL":
                    {
                        return reportTypes.CashRecieptsJournal;
                    }
                case "Current Sales Recap":
                    {
                        return reportTypes.CurrentSalesRecap;
                    }
                case "Purchase Analysis in Dollars":
                    {
                        return reportTypes.PurchaseAnalysisInDollars;
                    }
                case "Employee Master List":
                    {
                        return reportTypes.EmployeeMasterList;
                    }
                default:
                    {
                        break;  //return reportTypes.Other;
                    }
            }

            pointStart = firstSection2.IndexOf("Date:") + 16;   //  We just jumped past Time: xx:yy:zz
            pointEnd = firstSection2.IndexOf("Page", pointStart) - 1;


            ReportTitle = firstSection2.Substring(pointStart, pointEnd - pointStart).Trim();

            switch (ReportTitle)         //  this switch can't detect the daily reports unfortunately
            {
                case "Daily Sales Journal":
                    {
                        return reportTypes.DailySalesJournal;
                    }
                case "DAILY GENERAL LEDGER JOURNAL":
                    {
                        return reportTypes.CashRecieptsJournal;
                    }
                case "CASH RECEIPTS JOURNAL":
                    {
                        return reportTypes.CashRecieptsJournal;
                    }
                case "Sales History":
                    {
                        return reportTypes.SalesHistoryReport;
                    }
            }

            return reportTypes.Other;

        }




        private DateTime[] ReadDateSpan(string fil, reportTypes reportTypes)
        {
            DateTime[] dates = new DateTime[2];

            string date1 = "";
            string date2 = "";
            TextReader mt = new StreamReader(fil);

            int line = 99;   // the line number on which the date string appears
            int d1strt = 99;
            int d2strt = 99;
            int dateLen = 10;

            int countOfLines = GetLineCount(fil);

            string[] lines = new string[countOfLines];


            for (int i = 0; i < lines.Length - 1; i++)
            {
                lines[i] = mt.ReadLine();
            }
            // Lines of the report file are now indexed in lines[] and can be addressed as lines[countOfLines-1-2]

            switch (reportTypes)
            {
                case reportTypes.SalesHistoryReport:
                    {
                        line = 4;
                        d1strt = 26;
                        d2strt = 42;
                        break;
                    }
                case reportTypes.InventoryTransactionJournal:
                    {
                        line = 8;
                        d1strt = 73;
                        date1 = lines[line].Substring(d1strt, dateLen);
                        date2 = lines[countOfLines - 1 - 2].Substring(d1strt, dateLen);   // this is sloppy if last line doesn't have a date due to some row type phenomenon ...
                        dates[0] = Convert.ToDateTime(date1);
                        dates[1] = Convert.ToDateTime(date2);


                        return dates;
                    }
            }

            date1 = lines[line - 1].Substring(d1strt, dateLen);
            date2 = lines[line - 1].Substring(d2strt, dateLen);

            dates[0] = Convert.ToDateTime(date1);
            dates[1] = Convert.ToDateTime(date2);

            return dates;
        }

        public enum reportTypes { AccountsPayableLedger, InvoiceRegisterByEmployee, InvoiceRegisterByDay, InventoryTransactionJournal, MasterCustomerList, MasterProductList, ProductTxt, DailySalesJournal, CashRecieptsJournal, DailyGeneralLedgerJournal, Other, MasterVendorList, CheckBookRegister, MasterChartOfAccountsList, Unknown, MasterProductCategoryList, SalesHistoryReport, CurrentSalesRecap, PurchaseAnalysisInDollars, EmployeeMasterList };


        private static int GetLineCount(string fil)
        {
            StreamReader myReader = new StreamReader(fil);
            int countOfLines = 1;
            while (!myReader.EndOfStream)
            {
                countOfLines++;
                myReader.ReadLine();
            }
            return countOfLines;
        }

    }
}

