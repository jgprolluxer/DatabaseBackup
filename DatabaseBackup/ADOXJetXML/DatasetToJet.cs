using System;
using System.Collections.Generic;
using ADOX;
using ADODB;
using System.Data;
using System.IO;
using DataTypeEnum = ADOX.DataTypeEnum;

namespace ADOXJetXML
{
    public static class DatasetToJet
    {
        private static readonly List<string> ErrList = new List<string>();
        private static readonly List<string> ErrListParameters = new List<string>();

        public static void CopyDatasetSchemaToJetDb(string connectionString, DataSet ds, string filePath, ref string err)
        {
            var conn = new Connection();
            var cat = new Catalog();

            try
            {
                if (File.Exists(filePath)) File.Delete(filePath);
                cat.Create(connectionString);
                conn.Open(connectionString, "Admin");
                cat.ActiveConnection = conn;
                var errList2 = new List<string>();

                foreach (DataTable table in ds.Tables)
                {
                    errList2.Clear();

                    try
                    {
                        try
                        {
                            cat.Tables.Delete(table.TableName);
                        }
                        catch
                        {
                            // ignored
                        }
                        errList2.Add("TABLENAME: " + table.TableName);
                        var adoxTab = CopyDataTable(table, cat);
                        errList2.Add("COPYDATATABLE SUCCESS: " + table.TableName);
                        cat.Tables.Append(adoxTab);
                        errList2.Add("TABLEAPPEND SUCCESS: " + table.TableName);
                        MoveData(adoxTab, table);
                        errList2.Add("MOVEDATA SUCCESS: " + table.TableName);
                    }
                    catch (Exception e)
                    {
                        ErrList.AddRange(errList2);
                        ErrList.AddRange(ErrListParameters);
                        ErrList.Add("ERRORTABLE: " + table.TableName + " " + e.Message + " " + e.StackTrace);
                    }
                }
            }
            catch (Exception e)
            {
                err = e.Message + e.StackTrace;
            }
            finally
            {
                if (conn.State == (int)ConnectionState.Open)
                {
                    conn.Close();
                    //Marshal.ReleaseComObject(conn);
                }

                //Marshal.ReleaseComObject(cat);
                //GC.Collect();
                cat.ActiveConnection = null;
                ErrList.Add(err);
                err = string.Join(", ", ErrList.ToArray());
            }
        }

        private static Table CopyDataTable(DataTable table, Catalog cat)
        {
            var adoxTable = new Table
            {
                Name = table.TableName,
                ParentCatalog = cat
            };

            foreach (DataColumn col in table.Columns)
            {
                var adoxCol = new Column
                {
                    ParentCatalog = cat,
                    Name = col.ColumnName,
                    Type = TranslateDataTypeToAdoxDataType(col.DataType),
                    Attributes = ColumnAttributesEnum.adColNullable
                };

                if (col.MaxLength >= 0)
                    adoxCol.DefinedSize = col.MaxLength;

                adoxTable.Columns.Append(adoxCol, adoxCol.Type, adoxCol.DefinedSize);
            }

            return adoxTable;
        }

        private static DataTypeEnum TranslateDataTypeToAdoxDataType(Type type)
        {
            var guid = type.GUID.ToString();

            var adoxType =
                 guid == typeof(bool).GUID.ToString() ? DataTypeEnum.adBoolean :
                 guid == typeof(byte).GUID.ToString() ? DataTypeEnum.adUnsignedTinyInt :
                 guid == typeof(char).GUID.ToString() ? DataTypeEnum.adChar :
                 guid == typeof(DateTime).GUID.ToString() ? DataTypeEnum.adDate :
                 guid == typeof(decimal).GUID.ToString() ? DataTypeEnum.adDouble :
                 guid == typeof(double).GUID.ToString() ? DataTypeEnum.adDouble :
                 guid == typeof(short).GUID.ToString() ? DataTypeEnum.adSmallInt :
                 guid == typeof(int).GUID.ToString() ? DataTypeEnum.adInteger :
                 guid == typeof(long).GUID.ToString() ? DataTypeEnum.adBigInt :
                 guid == typeof(sbyte).GUID.ToString() ? DataTypeEnum.adTinyInt :
                 guid == typeof(float).GUID.ToString() ? DataTypeEnum.adSingle :
                 guid == typeof(string).GUID.ToString() ? DataTypeEnum.adLongVarWChar :
                 guid == typeof(TimeSpan).GUID.ToString() ? DataTypeEnum.adDouble :
                 guid == typeof(ushort).GUID.ToString() ? DataTypeEnum.adSmallInt :
                 guid == typeof(uint).GUID.ToString() ? DataTypeEnum.adInteger :
                 guid == typeof(ulong).GUID.ToString() ? DataTypeEnum.adBigInt :
                 DataTypeEnum.adBinary;

            return adoxType;
        }

        private static Command AdoxTableInsertCommand(DataTable aTable, ref List<DataTypeEnum> aType)
        {
            var result = new Command
            {
                CommandText = string.Format("INSERT INTO {0} ({1}) values({2}) ", aTable.TableName, "{0}", "{1}")
            };
            var colNames = string.Empty;
            var colVals = string.Empty;

            for (var i = 0; i < aTable.Columns.Count; i++)
            {
                var adoxCol = aTable.Columns[i];
                var name = adoxCol.ToString();
                var type = TranslateDataTypeToAdoxDataType(adoxCol.DataType);
                aType.Add(type);
                ErrListParameters.Add("COLUMN/TYPE: " + name + "/" + type);
                switch (type)
                {
                    case DataTypeEnum.adVarBinary: break;
                    default:
                        colNames += (colNames != string.Empty ? "," : "") + name;
                        if (type == DataTypeEnum.adInteger)
                        {
                            colVals += (colVals != string.Empty ? "," : "") + "?";
                        }
                        else
                        {
                            colVals += (colVals != string.Empty ? "," : "") + "'?'";
                        }
                        break;
                }
            }

            result.CommandText = string.Format(result.CommandText, colNames, colVals);
            ErrListParameters.Add(result.CommandText);
            return result;
        }

        private static void MoveData(Table adoxTab, DataTable aTable)
        {
            ErrListParameters.Clear();
            var aType = new List<DataTypeEnum>();
            var cmd = AdoxTableInsertCommand(aTable, ref aType); //adoxTab);
            cmd.ActiveConnection = (Connection) adoxTab.ParentCatalog.ActiveConnection;

            foreach (DataRow row in aTable.Rows)
            {
                //var arry = row.ItemArray;
                var a = new List<string>();
                var i = 0;

                foreach (var item in row.ItemArray)
                {
                    switch (aType[i])
                    {
                        case DataTypeEnum.adVarBinary: break;
                        case DataTypeEnum.adDate:
                            ErrListParameters.Add("DATE 1: (" + item + ") ");
                            ErrListParameters.Add("DATE 2: (" + GetDate(item) + ") ");
                            a.Add(GetDate(item));
                            break;
                        case DataTypeEnum.adBoolean:
                            ErrListParameters.Add("BOOLEAN 1: (" + item + ") ");
                            ErrListParameters.Add("BOOLEAN 2: (" + GetBoolean(item) + ") ");
                            a.Add(GetBoolean(item));
                            break;
                        default:
                            ErrListParameters.Add("STRING: (" + item + ") ");
                            a.Add(item.ToString());
                            break;
                    }

                    i++;
                }

                ErrListParameters.Add("PARAMETERS: (" + string.Join(", ", a.ToArray()) + ") ");
                object arry = a.ToArray();
                object count;
                cmd.Execute(out count, ref arry, 1);
            }

            ErrListParameters.Clear();
        }

        //private static int GetInt(object o)
        //{
        //    int r;
        //    int.TryParse((string)o, out r);
        //    return r;
        //}

        private static string GetDate(object o)
        {
            DateTime r;
            if(!DateTime.TryParse(o.ToString(), out r))
                r = new DateTime(); //todo fix this
            return r.ToString("dd/MM/yyyy HH:mm:ss");
        }

        private static string GetBoolean(object o)
        {
            bool r;
            bool.TryParse((string)o, out r);
            return r.ToString();
        }
    }
}
