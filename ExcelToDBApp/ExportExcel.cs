using ExcelDataReader;
using ExcelToDBAppDBLib.DAL;
using ExcelToDBAppModel.HR;
using ExcelToDBAppModel.HR.Designation;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace ExcelToDBApp
{
    public partial class ExportExcel : Form
    {
        public ExportExcel()
        {
            InitializeComponent();
        }
        private readonly ImportData objID = new ImportData();
        readonly AdoHelper adoHelper = new AdoHelper();
        DataTableCollection tableCollection;
        DataSet result;
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.Xlsx|All files (*.*)|*.*" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFileName.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            cboSheet.Items.Clear();
                            foreach (DataTable dataTable in tableCollection)
                                cboSheet.Items.Add(dataTable.TableName);
                        }
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dataTable = tableCollection[cboSheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dataTable;
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                if(result == null)
                {
                    MessageBox.Show("No file is selected", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                for (int i = 0; i < result.Tables.Count; i++)
                {
                    DataTable dt = result.Tables[i];
                    if (adoHelper.isDataTableContainRecords(dt))
                    {
                        if (tableCollection[i].TableName == "Department")
                        {
                            List<HR_Department_VM> List = new List<HR_Department_VM>();

                            List = adoHelper.ConvertDataTable<HR_Department_VM>(dt);
                            if (List != null)
                            {
                                for (int j = 0; j < List.Count; j++)
                                {
                                    objID.ImportHRDepartment(List[j]);
                                }
                            }
                        }
                        if (tableCollection[i].TableName == "Designation")
                        {
                            List<HRDesignation_VM> HRDesignation_List = new List<HRDesignation_VM>();

                            HRDesignation_List = adoHelper.ConvertDataTable<HRDesignation_VM>(dt);
                            //for (int k = 0; k < dt.Rows.Count; k++)
                            //{
                            //    HRDesignation_VM hRDesignation_VM = new HRDesignation_VM
                            //    {
                            //        DesignationId = Convert.ToInt32(dt.Rows[k]["DesignationId"]),
                            //        DesignationName = dt.Rows[k]["DesignationName"].ToString(),
                            //        CreatedBy = Convert.ToInt32(dt.Rows[k]["CreatedBy"])
                            //    };
                            //    HRDesignation_List.Add(hRDesignation_VM);
                            //}
                            if (HRDesignation_List != null)
                            {
                                for (int j = 0; j < HRDesignation_List.Count; j++)
                                {
                                    objID.ImportHRDesignation(HRDesignation_List[j]);
                                }
                            }
                        }
                    }
                }
                MessageBox.Show("Data Imported successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
