using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aplicativo
{
    public class formEst
    {
        Application SBO_Application_3;
        SAPbobsCOM.Company company;

        public formEst(Application app, SAPbobsCOM.Company comp)
        {
            SBO_Application_3 = app;
            company = comp;
        }
        public void ShowForm()
        {
            try
            {
                FormCreationParams formCreationParams = (FormCreationParams)SBO_Application_3.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                formCreationParams.XmlData = Properties.Resources.TransferenciaEstoque;
                formCreationParams.UniqueID = "TransEst" + Guid.NewGuid().ToString().Substring(0,10);
                formCreationParams.FormType = "est";
                Form est = SBO_Application_3.Forms.AddEx(formCreationParams);
                est.Visible = true;

                
                Grid grid = (Grid)est.Items.Item("grid").Specific;
                grid.AutoResizeColumns();

                SBO_Application_3.StatusBar.SetText("Tela iniciada", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex) 
            {
                SBO_Application_3.MessageBox(ex.Message);
            }
        }


        public void itemEventEstoq(ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormTypeEx == "est" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_FORM_LOAD)
            {
                var oForm = SBO_Application_3.Forms.Item(pVal.FormUID);

            }

            if (pVal.FormTypeEx == "est" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnCar1")
            {
                Grid(pVal);
            }


            if (pVal.FormTypeEx == "est" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btn_Merc")
            {
                AdicionarMercadoria(pVal);
            }
        }


        public void Grid(ItemEvent pval)
        {
            var oForm = SBO_Application_3.Forms.Item(pval.FormUID);
            try
            {
                oForm.Freeze(true);
                EditText item1 = (EditText)oForm.Items.Item("Item1").Specific;

                DataTable tb = (DataTable)oForm.DataSources.DataTables.Item("DTItem1");

                Recordset ds = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = @"SELECT * FROM OITW WHERE ""ItemCode"" = '" + item1.Value + "'";
                ds.DoQuery(query);

                tb.Rows.Clear();

                if (ds.RecordCount > 0)
                {

                    while (!ds.EoF)
                    {
                        tb.Rows.Add();
                        string itemCode = ds.Fields.Item("itemCode").Value.ToString();
                        tb.SetValue("itemCode", tb.Rows.Count - 1, itemCode);

                        string whsCode = ds.Fields.Item("whsCode").Value.ToString();
                        tb.SetValue("whsCode", tb.Rows.Count - 1, whsCode);

                        string OnHand = ds.Fields.Item("OnHand").Value.ToString();
                        tb.SetValue("OnHand", tb.Rows.Count - 1, OnHand);

                        string IsCommited = ds.Fields.Item("IsCommited").Value.ToString();
                        tb.SetValue("IsCommited", tb.Rows.Count - 1, IsCommited);

                        string OnOrder = ds.Fields.Item("OnOrder").Value.ToString();
                        tb.SetValue("OnOrder", tb.Rows.Count - 1, OnOrder);

                        ds.MoveNext();
                    }

                }

               

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false); 
                SBO_Application_3.MessageBox(ex.Message);
                
            }
        }


        public void AdicionarMercadoria(ItemEvent pval)
        {
            try
            {
                //Organizar os valores em variavéis
                var oForm = SBO_Application_3.Forms.Item(pval.FormUID);
                EditText it = (EditText)oForm.Items.Item("Item").Specific;
                EditText quant = (EditText)oForm.Items.Item("Quant").Specific;
                EditText dep = (EditText)oForm.Items.Item("Dep").Specific;

                string item = it.Value;
                string quantidade = quant.Value;
                string deposito = dep.Value;

                //Validação dos campos
                if (string.IsNullOrEmpty(item))
                {
                    SBO_Application_3.MessageBox("Preencha o campo Item");
                    return;
                }


                if (string.IsNullOrEmpty(quantidade))
                {
                    SBO_Application_3.MessageBox("Preencha o campo Quantidade");
                    return;
                }


                if (string.IsNullOrEmpty(deposito))
                {
                    SBO_Application_3.MessageBox("Preencha o campo Deposito");
                    return;
                }

                //Insert dos dados em documento SAP
                Documents doc = (Documents)company.GetBusinessObject(BoObjectTypes.oInventoryGenEntry);
                doc.Comments = "Observação";
                doc.Lines.ItemCode = item;
                doc.Lines.Quantity = Convert.ToDouble(quantidade);
                doc.Lines.WarehouseCode = deposito;
                int result = doc.Add();



                //Validação para captar mensagem em caso de erro
                if(result != 0)
                {
                    int codigo;
                    string messag;
                    company.GetLastError(out codigo, out messag);   
                    throw new Exception("Erro ao realizar entrada de mercadoria. " + codigo + messag);  //Mostra o erro


                }
                else
                {
                    SBO_Application_3.MessageBox("Entrada realizada com sucesso");
                }

            }
            catch(Exception ex)
            {
                SBO_Application_3.MessageBox(ex.Message);
            }

        }

    }
}
