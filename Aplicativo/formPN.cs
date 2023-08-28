using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;


namespace Aplicativo.Forms
{
    class formPN //Classe que controla a tela 
    {
        Application SBO_Application_2;
        SAPbobsCOM.Company company;
        public formPN(Application app, SAPbobsCOM.Company comp)
        {
            SBO_Application_2 = app;
            company = comp;
        }
        public void ShowForm()
        {
            try
            {
                FormCreationParams oCreationParams = (FormCreationParams)SBO_Application_2.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                oCreationParams.XmlData = Properties.Resources.Pn;
                oCreationParams.UniqueID = "Pn" + Guid.NewGuid().ToString().Substring(0, 10);
                oCreationParams.FormType = "Pn";
                Form form = SBO_Application_2.Forms.AddEx(oCreationParams);
                form.Visible = true;


                Grid grid = (Grid)form.Items.Item("grdPNs").Specific;
                grid.AutoResizeColumns();


                SBO_Application_2.StatusBar.SetText("Tela iniciada", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public void itemEventPn(ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            

            if (pVal.FormTypeEx == "Pn" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnCar")
            {
                Form oForm = SBO_Application_2.Forms.Item(pVal.FormUID);
                ListaPns(pVal);
            }

            if (pVal.FormTypeEx == "Pn" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "OpAdd")
            {
                Form oForm = SBO_Application_2.Forms.Item(pVal.FormUID);
                oForm.PaneLevel = 3;

                CarregComboBox(pVal);

                Limpardados(pVal);

            }

            if (pVal.FormTypeEx == "Pn" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "OpMod")
            {
                Form oForm = SBO_Application_2.Forms.Item(pVal.FormUID);
                oForm.PaneLevel = 4;

                CarregComboBox(pVal);
                
                Limpardados(pVal);

            }

            if (pVal.FormTypeEx == "Pn" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == "combGrup")
            {
                Form oForm = SBO_Application_2.Forms.Item(pVal.FormUID);
                ComboBox GrupoPn = (ComboBox)oForm.Items.Item("combGrup").Specific;
                ComboBox TipoPn = (ComboBox)oForm.Items.Item("combTip").Specific;


                if (GrupoPn.Selected.Description == "Definir novo")
                {
                    if (TipoPn.Value == "F")
                    {
                        //ABRE DOCUMENTO
                        SBO_Application_2.ActivateMenuItem("10753");

                        //PEGA O FORMULÁRIO DO DOCUMENTO ABERTO
                        //SAPbouiCOM.Form FormDocumento;
                        //FormDocumento = SBO_Application_2.Forms.ActiveForm;

                        ////MUDA MODO FORMULÁRIO PARA >> PESQUISA
                        //FormDocumento.Mode = BoFormMode.fm_FIND_MODE;

                        ////RECUPERA ITENS DO FORMULÁRIO
                        //Item txtPesquisa = (Item)FormDocumento.Items.Item(5);
                        //EditText edtPesquisa = (EditText)txtPesquisa.Specific;
                        //edtPesquisa.Value = docentry

                    //SIMULA CLICK PESQUISAR (btnPrincipal)
                    //    Item btnPrincipal = (Item)FormDocumento.Items.Item("1");
                    //btnPrincipal.Click(BoCellClickType.ct_Regular);
                    }

                    else
                    {
                        SBO_Application_2.ActivateMenuItem("10754");
                    }

                }

            }

            if (pVal.FormTypeEx == "Pn" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == "combMoed")
            {
                Form oForm = SBO_Application_2.Forms.Item(pVal.FormUID);
                ComboBox GrupoPn = (ComboBox)oForm.Items.Item("combGrup").Specific;
                ComboBox TipoPn = (ComboBox)oForm.Items.Item("combTip").Specific;


                if (GrupoPn.Selected.Description == "Definir novo")
                {
                    
                   SBO_Application_2.ActivateMenuItem("8450");
                    

                }

            }

            if (pVal.FormTypeEx == "Pn" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_CLICK && (pVal.ItemUID == "btnAdd" || pVal.ItemUID == "btnAt"))
            {
                Form oForm = SBO_Application_2.Forms.Item(pVal.FormUID);
                AdicionarAtualizar(pVal);
            }
            
            if (pVal.FormTypeEx == "Pn" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnPesq")
            {
                Form oForm = SBO_Application_2.Forms.Item(pVal.FormUID);
                PesquisarPn(pVal);
            }

            if (pVal.FormTypeEx == "Pn" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "Validacao")
            {
                Form oForm = SBO_Application_2.Forms.Item(pVal.FormUID);
                ComboBox tipo = (ComboBox)oForm.Items.Item("combTip").Specific;
                if (tipo.Value == "C" || tipo.Value == "L")
                {
                    oForm.PaneLevel = 5;
                }
                else
                {
                    oForm.PaneLevel = 6;
                }
            }

            if (pVal.FormTypeEx == "Pn" && !pVal.BeforeAction && pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnVal")
            {
                Form oForm = SBO_Application_2.Forms.Item(pVal.FormUID);
                Validacao(pVal);
            }

        }
        
        public void ListaPns(ItemEvent pval)
        {

            Form oForm = SBO_Application_2.Forms.Item(pval.FormUID);
            ComboBox tipo = (ComboBox)oForm.Items.Item("combTip").Specific;
            try
            {

                oForm.Freeze(true);

                DataTable tb = (DataTable)oForm.DataSources.DataTables.Item("TbPn");
                Recordset ds2 = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

                //Consulta
                if (tipo.Value == "C")
                {
                    string query = @"SELECT TOP 10 CardCode, CardName, CardType, Balance, E_Mail FROM OCRD WHERE CardType = 'C' ORDER BY CardType DESC";
                    ds2.DoQuery(query);
                }

                else if ((tipo.Value == "S"))
                {
                    string query = @"SELECT TOP 10 CardCode, CardName, CardType, Balance, E_Mail FROM OCRD WHERE CardType = 'S' ORDER BY CardType DESC";
                    ds2.DoQuery(query);
                }
                else
                {
                    string query = @"SELECT TOP 10 CardCode, CardName, CardType, Balance, E_Mail FROM OCRD WHERE CardType = 'L' ORDER BY CardType DESC";
                    ds2.DoQuery(query);
                }

                //Carregar dados no DataTable
                tb.Rows.Clear();

                if (ds2.RecordCount > 0)
                {

                    while (!ds2.EoF)
                    {
                        tb.Rows.Add();

                        string CardCode = ds2.Fields.Item("CardCode").Value.ToString();
                        tb.SetValue("CardCode", tb.Rows.Count - 1, CardCode);

                        string CardName = ds2.Fields.Item("CardName").Value.ToString();
                        tb.SetValue("CardName", tb.Rows.Count - 1, CardName);

                        string CardType = ds2.Fields.Item("CardType").Value.ToString();
                        tb.SetValue("CardType", tb.Rows.Count - 1, CardType);

                        string Balance = ds2.Fields.Item("Balance").Value.ToString();
                        if (!string.IsNullOrEmpty(Balance))
                        {
                            double balance = Convert.ToDouble(Balance);
                            tb.SetValue("Balance", tb.Rows.Count - 1, balance);
                        }

                        string E_Mail = ds2.Fields.Item("E_Mail").Value.ToString();
                        tb.SetValue("E_Mail", tb.Rows.Count - 1, E_Mail);

                        ds2.MoveNext();
                    }
                }

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                SBO_Application_2.MessageBox("Erro ao carregar Dados na Grid. Erro: " + ex.Message);
            }
        }

        public void AdicionarAtualizar(ItemEvent pval)

        {
            Form oForm = SBO_Application_2.Forms.Item(pval.FormUID);

            #region Variavéis
            //Organizar os valores em variavéis
            EditText codPn = (EditText)oForm.Items.Item("Cod").Specific;
            EditText NomePn = (EditText)oForm.Items.Item("Nome").Specific;
            ComboBox GrupoPn = (ComboBox)oForm.Items.Item("combGrup").Specific;
            ComboBox MoedaPn = (ComboBox)oForm.Items.Item("combMoed").Specific;
            EditText tel1Pn = (EditText)oForm.Items.Item("tel1").Specific;
            EditText tel2Pn = (EditText)oForm.Items.Item("tel2").Specific;
            EditText telCePn = (EditText)oForm.Items.Item("telCe").Specific;
            EditText emailPn = (EditText)oForm.Items.Item("email").Specific;
            EditText NomeFPn = (EditText)oForm.Items.Item("NomF").Specific;

            string cod = codPn.Value;
            string Nome = NomePn.Value;
            string Grupo = GrupoPn.Value;
            string Moeda = MoedaPn.Value;
            string tel1 = tel1Pn.Value;
            string tel2 = tel2Pn.Value;
            string telCe = telCePn.Value;
            string email = emailPn.Value;
            string NomeF = NomeFPn.Value;
            #endregion

            #region Adicionar cadastro
            try
            {
                if (pval.ItemUID == "btnAdd")
                {
                    //Insert dos dados no cadastro de PN
                    BusinessPartners doc = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                    doc.CardCode = cod;
                    doc.CardName = Nome;
                    doc.GroupCode = int.Parse(Grupo); 
                    doc.Currency = Moeda;   
                    doc.Phone1 = tel1;
                    doc.Phone2 = tel2;
                    doc.Cellular = telCe;
                    doc.FreeText = "Cadastro inserido através da tela Parceiro de negócio do Addon Teste";

                    int result = doc.Add();

                    //Validação para captar mensagem em caso de erro
                    if (result != 0)
                    {
                        int codigo;
                        string messag;
                        company.GetLastError(out codigo, out messag);
                        throw new Exception("Erro ao realizar Cadastro do PN. " + codigo + messag);  //Mostra o erro
                    }
                    else
                    {
                        SBO_Application_2.MessageBox("Cadastro realizado com sucesso");
                    }
                    
                    Limpardados(pval);
                }

            }
            catch (Exception ex)
            {
                SBO_Application_2.MessageBox("Erro ao adicionar dados do PN" + ex.Message);
            }
            #endregion

            #region Modificar registro
            try
            {
                if (pval.ItemUID == "btnAt")
                {
                    //Update dos dados no cadastro de PN
                    BusinessPartners doc = (BusinessPartners)company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                    doc.GetByKey(cod);
                    doc.CardName = Nome;
                    doc.GroupCode = int.Parse(Grupo);
                    doc.Currency = Moeda;   
                    doc.Phone1 = tel1;
                    doc.Phone2 = tel2;
                    doc.Cellular = telCe;
                    doc.FreeText = "Cadastro modificado através da tela Parceiro de negócio do Addon Teste";

                    int result2 = doc.Update();

                    //Validação para captar mensagem em caso de erro
                    if (result2 != 0)
                    {
                        int codigo;
                        string messag;
                        company.GetLastError(out codigo, out messag);
                        throw new Exception("Erro ao realizar modificação no cadastro do PN. " + codigo + messag);  //Mostra o erro
                    }
                    else
                    {
                        SBO_Application_2.MessageBox("Cadastro atualizado com sucesso");
                    }
                   
                }
                Limpardados(pval);

            }
            catch (Exception ex)
            {
                SBO_Application_2.MessageBox("Erro ao modificar os dados do PN" + ex.Message);
            }
            #endregion
        }

        public void PesquisarPn(ItemEvent pval)
        {
           

            Form oForm = SBO_Application_2.Forms.Item(pval.FormUID);


            #region Variavéis
            //Organizar os valores em variavéis
            EditText codPn = (EditText)oForm.Items.Item("Cod").Specific;
            EditText NomePn = (EditText)oForm.Items.Item("Nome").Specific;
            ComboBox GrupoPn = (ComboBox)oForm.Items.Item("combGrup").Specific;
            ComboBox MoedaPn = (ComboBox)oForm.Items.Item("combMoed").Specific;
            EditText tel1Pn = (EditText)oForm.Items.Item("tel1").Specific;
            EditText tel2Pn = (EditText)oForm.Items.Item("tel2").Specific;
            EditText telCePn = (EditText)oForm.Items.Item("telCe").Specific;
            EditText emailPn = (EditText)oForm.Items.Item("email").Specific;
            EditText NomeFPn = (EditText)oForm.Items.Item("NomF").Specific;

            string cod = codPn.Value;

            #endregion

            //Consulta
            Recordset ds = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = @"select * from OCRD Inner Join OCRG on OCRD.GroupCode = OCRG.GroupCode WHERE CardCode ='" + cod + "'";
            ds.DoQuery(query);

            //Trazer valores 
            NomePn.Value = ds.Fields.Item("CardName").Value.ToString();
            GrupoPn.Select(ds.Fields.Item("GroupCode").Value.ToString());
            MoedaPn.Select(ds.Fields.Item("Currency").Value.ToString());;
            tel1Pn.Value = ds.Fields.Item("Phone1").Value.ToString();
            tel2Pn.Value = ds.Fields.Item("Phone2").Value.ToString();
            telCePn.Value = ds.Fields.Item("Cellular").Value.ToString();
            emailPn.Value = ds.Fields.Item("E_Mail").Value.ToString();
            NomeFPn.Value = ds.Fields.Item("AliasName").Value.ToString();

        }

        public void Limpardados(ItemEvent pval)
        {
            Form oForm = SBO_Application_2.Forms.Item(pval.FormUID);
            oForm.Freeze(true);

            //Campos
            EditText codPn = (EditText)oForm.Items.Item("Cod").Specific;
            EditText NomePn = (EditText)oForm.Items.Item("Nome").Specific;
            ComboBox GrupoPn = (ComboBox)oForm.Items.Item("combGrup").Specific;
            ComboBox MoedaPn = (ComboBox)oForm.Items.Item("combMoed").Specific;
            EditText tel1Pn = (EditText)oForm.Items.Item("tel1").Specific;
            EditText tel2Pn = (EditText)oForm.Items.Item("tel2").Specific;
            EditText telCePn = (EditText)oForm.Items.Item("telCe").Specific;
            EditText emailPn = (EditText)oForm.Items.Item("email").Specific;
            EditText NomeFPn = (EditText)oForm.Items.Item("NomF").Specific;

            codPn.Value = "";
            NomePn.Value = "";
            GrupoPn.Select("0");
            MoedaPn.Select("0");
            tel1Pn.Value = "";
            tel2Pn.Value = "";
            telCePn.Value = "";
            emailPn.Value = "";
            NomeFPn.Value = "";

            oForm.Freeze(false);
        }

        public void CarregComboBox(ItemEvent pval) 
        {
            
            Form oForm = SBO_Application_2.Forms.Item(pval.FormUID);
            oForm.Freeze(true);

            try
            {
                //ComboBox Grupos
                ComboBox GrupoPn = (ComboBox)oForm.Items.Item("combGrup").Specific;

                SAPbobsCOM.Recordset ds = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                while (GrupoPn.ValidValues.Count > 0)
                {
                    GrupoPn.ValidValues.Remove(0, BoSearchKey.psk_Index);
                }

                string query = @"select * from OCRG";
                ds.DoQuery(query);
                while (!ds.EoF)
                {
                    GrupoPn.ValidValues.Add(ds.Fields.Item("GroupCode").Value.ToString(), ds.Fields.Item("GroupName").Value.ToString());
                    ds.MoveNext();
                }
                
                GrupoPn.ValidValues.Add("0", "");
                GrupoPn.ValidValues.Add("", "Definir novo");


                //ComboBox Moeda
                ComboBox MoedaPn = (ComboBox)oForm.Items.Item("combMoed").Specific;

                while (MoedaPn.ValidValues.Count > 0)
                {
                    MoedaPn.ValidValues.Remove(0, BoSearchKey.psk_Index);
                }

                string query2 = @"select * from OCRN";
                ds.DoQuery(query2);
                while (!ds.EoF)
                {
                    MoedaPn.ValidValues.Add(ds.Fields.Item("CurrCode").Value.ToString(), ds.Fields.Item("CurrName").Value.ToString());
                    ds.MoveNext();
                }

                MoedaPn.ValidValues.Add("0", "");
                MoedaPn.ValidValues.Add("", "Definir novo");
            }
            catch (Exception ex)
            {
                SBO_Application_2.MessageBox("Erro ao alimentar comboBox:" + ex);
            }
            oForm.Freeze(false);
        }

        public void Validacao(ItemEvent pval)
        {
            try
            {
                #region Cliente

                //Atribuir o Id do cliente em  uma variavel
                Form frmPN = SBO_Application_2.Forms.Item(pval.FormUID);
                EditText IdC = (EditText)frmPN.Items.Item("IdC").Specific;
                string CodCl = IdC.Value;

                if (!string.IsNullOrEmpty(CodCl))
                {

                    //Trazer valor para o Nome do Pn
                    Recordset ds = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    ds.DoQuery(@"SELECT ""CardName"" FROM OCRD WHERE ""CardCode"" = '" + CodCl + "' ");
                    EditText Cliente = (EditText)frmPN.Items.Item("Cliente").Specific;

                    Cliente.Value = ds.Fields.Item("CardName").Value.ToString();
                }

                #endregion

                #region Fornecedor

                //Atribuir o Id do fornecedor em  uma variavel
                EditText IdF = (EditText)frmPN.Items.Item("IdF").Specific;
                string CodFn = IdF.Value;

                if (!string.IsNullOrEmpty(CodFn))
                {

                    //Trazer valor para o Nome do Pn
                    Recordset ds1 = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    ds1.DoQuery(@"SELECT ""CardName"" FROM OCRD WHERE ""CardCode"" = '" + CodFn + "' ");
                    EditText Forn = (EditText)frmPN.Items.Item("Forn").Specific;

                    Forn.Value = ds1.Fields.Item("CardName").Value.ToString();
                }
                #endregion
            }
            catch (Exception ex)
            {
                SBO_Application_2.MessageBox("Erro ao realizar consulta. Erro: " + ex.Message);
            }
        }

    }
}
