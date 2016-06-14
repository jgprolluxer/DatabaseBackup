using System;
using ADOX;
using ADODB;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;

namespace ADOXJetXML
{
    public static class DatasetToJet
    {
        public static void CopyDatasetSchemaToJetDb(string connectionString, DataSet ds, string filePath)
        {
            Connection conn = new ConnectionClass();
            Catalog cat = new CatalogClass();

            try
            {
                if (File.Exists(filePath)) File.Delete(filePath);
                cat.Create(connectionString);
                conn.Open(connectionString, "Admin");
                cat.ActiveConnection = conn;

                foreach (DataTable table in ds.Tables)
                {
                    try
                    {
                        cat.Tables.Delete(table.TableName);
                    }
                    catch
                    {
                        // ignored
                    }

                    var adoxTab = CopyDataTable(table, cat);
                    cat.Tables.Append(adoxTab);
                    MoveData(adoxTab, table);
                }
            }
            catch (Exception)
            {
                //
            }
            finally
            {
                if (conn.State == (int)ConnectionState.Open)
                {
                    conn.Close();
                    Marshal.ReleaseComObject(conn);
                }

                Marshal.ReleaseComObject(cat);
                GC.Collect();
                cat.ActiveConnection = null;
            }
        }

        private static Table CopyDataTable(DataTable table, Catalog cat)
        {
            Table adoxTable = new TableClass();
            adoxTable.Name = table.TableName;
            adoxTable.ParentCatalog = cat;

            foreach (DataColumn col in table.Columns)
            {
                Column adoxCol = new ColumnClass();
                adoxCol.ParentCatalog = cat;
                adoxCol.Name = col.ColumnName;
                adoxCol.Type = TranslateDataTypeToAdoxDataType(col.DataType);
                adoxCol.Attributes = ColumnAttributesEnum.adColNullable;

                if (col.MaxLength >= 0)
                    adoxCol.DefinedSize = col.MaxLength;

                adoxTable.Columns.Append(adoxCol, adoxCol.Type, adoxCol.DefinedSize);
            }

            return adoxTable;
        }

        private static ADOX.DataTypeEnum TranslateDataTypeToAdoxDataType(Type type)
        {
            var guid = type.GUID.ToString();
            var adoxType =
             guid == typeof(bool).GUID.ToString() ? ADOX.DataTypeEnum.adBoolean :
             guid == typeof(byte).GUID.ToString() ? ADOX.DataTypeEnum.adUnsignedTinyInt :
             guid == typeof(char).GUID.ToString() ? ADOX.DataTypeEnum.adChar :
             guid == typeof(DateTime).GUID.ToString() ? ADOX.DataTypeEnum.adDate :
             guid == typeof(decimal).GUID.ToString() ? ADOX.DataTypeEnum.adDouble :
             guid == typeof(double).GUID.ToString() ? ADOX.DataTypeEnum.adDouble :
             guid == typeof(short).GUID.ToString() ? ADOX.DataTypeEnum.adSmallInt :
             guid == typeof(int).GUID.ToString() ? ADOX.DataTypeEnum.adInteger :
             guid == typeof(long).GUID.ToString() ? ADOX.DataTypeEnum.adBigInt :
             guid == typeof(sbyte).GUID.ToString() ? ADOX.DataTypeEnum.adTinyInt :
             guid == typeof(float).GUID.ToString() ? ADOX.DataTypeEnum.adSingle :
             guid == typeof(string).GUID.ToString() ? ADOX.DataTypeEnum.adLongVarWChar :
             guid == typeof(TimeSpan).GUID.ToString() ? ADOX.DataTypeEnum.adDouble :
             guid == typeof(ushort).GUID.ToString() ? ADOX.DataTypeEnum.adUnsignedSmallInt :
             guid == typeof(uint).GUID.ToString() ? ADOX.DataTypeEnum.adUnsignedInt :
             guid == typeof(ulong).GUID.ToString() ? ADOX.DataTypeEnum.adUnsignedBigInt :
             ADOX.DataTypeEnum.adVarBinary;
            return adoxType;
        }

        private static Command AdoxTableInsertCommand(Table adoxTab)
        {
            Command result = new CommandClass();
            var conn = adoxTab.ParentCatalog.ActiveConnection;
            result.ActiveConnection = (ConnectionClass)conn;
            result.CommandText = string.Format("INSERT INTO {0} ({1}) values({2}) ", adoxTab.Name, "{0}", "{1}");
            var colNames = string.Empty;
            var colVals = string.Empty;

            for (var i = 0; i < adoxTab.Columns.Count; i++)
            {
                var adoxCol = adoxTab.Columns[i];
                var name = adoxCol.Name;
                var type = adoxCol.Type;

                switch (type)
                {
                    case ADOX.DataTypeEnum.adVarBinary: break;
                    default:
                        colNames += (colNames != string.Empty ? "," : "") + name;
                        colVals += (colVals != string.Empty ? "," : "") + "?";
                        break;
                }
            }
            
            result.CommandText = string.Format(result.CommandText, colNames, colVals);

            return result;
        }

        private static void MoveData(Table adoxTab, DataTable aTable)
        {
            var cmd = AdoxTableInsertCommand(adoxTab);

            foreach (DataRow row in aTable.Rows)
            {
                object arry = row.ItemArray;
                object i;
                cmd.Execute(out i, ref arry, 1);
            }
        }
    }
}
