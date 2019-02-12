using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Logica;
using System.Timers;

namespace Comunicador
{
    public partial class Comunica : Form
    {
        public string[] sFeriados = new string[10];
        public Comunica()
        {
            InitializeComponent();
        }

        private int ColumnWith(DataGridView _dtGrid, double _dColWith)
        {

            double dW = _dtGrid.Width - 10;
            double dTam = _dColWith;
            double dPor = dTam / 100;
            dTam = dW * dPor;
            dTam = Math.Truncate(dTam);

            return Convert.ToInt32(dTam);
        }

        private void Comunica_Load(object sender, EventArgs e)
        {
            tssVersion.Text += "1.0.1.25";

            sFeriados[0] = "05/02/2018";
            sFeriados[1] = "19/03/2018";
            sFeriados[2] = "30/03/2018";
            sFeriados[3] = "30/04/2018";
            sFeriados[4] = "19/11/2018";
            sFeriados[5] = "25/12/2018";

            timer1.Start();
        }

        #region regAlertasServicio
        private void AlertaProductividad()
        {
            ComunicaLogica com = new ComunicaLogica();
            com.Proceso = "C100";
            com.Referencia = "PRODTECDIA";
            if (!ComunicaLogica.AlertaDiaria(com))
                ComunicaLogica.AlertaPRODUCTEC();
        }
        private void AlertaMaqCosto()
        {
            ComunicaLogica com = new ComunicaLogica();
            com.Proceso = "B190";
            com.Referencia = "MAQCTOSUP";
            if (!ComunicaLogica.AlertaMensual(com))
                ComunicaLogica.AlertaMaqCosto();
        }
        private void AlertaTOP10Servicios()
        {
            ComunicaLogica com = new ComunicaLogica();
            com.Proceso = "C060";
            com.Referencia = "TOP10SERV";
            if (!ComunicaLogica.AlertaMensual(com))
                ComunicaLogica.AlertaTop10Servicios();
        }
        private void AlertaTOP10CtoServ()
        {
            ComunicaLogica com = new ComunicaLogica();
            com.Proceso = "C060";
            com.Referencia = "TOP10CTOSE";
            if (!ComunicaLogica.AlertaMensual(com))
                ComunicaLogica.AlertaTop10CtoServ();
        }
        private void AlertaOSFECOM()
        {
            ComunicaLogica com = new ComunicaLogica();
            com.Proceso = "C060";
            com.Referencia = "OSFECOM";
            if(!ComunicaLogica.AlertaDiaria(com))
                ComunicaLogica.AlertaOSFECOM();
            //_lbAlerFCom = false;
        }
        private void AlertaFechaPromesa()
        {
            ComunicaLogica com = new ComunicaLogica();
            com.Proceso = "C060";
            com.Referencia = "OSPROCOM";
            if (!ComunicaLogica.AlertaDiaria(com))
                ComunicaLogica.AlertaOSPROCOM();
        }
        private void AlertaExistencia()
        {
            ComunicaLogica com = new ComunicaLogica();
            com.Proceso = "B220";
            com.Referencia = "PTEEXISMIN";
            if (!ComunicaLogica.AlertaDiaria(com))
                ComunicaLogica.AlertaEXIS();
        }
        #endregion

        #region regAlertasProd
        private void AlertaPreviaSetUp()
        {
            ComunicaCPROLogica.AlertaPRESETUP();
        }

        private void AlertaKanban()
        {
            ComunicaCPROLogica.AlertaKANBAN();
        }
        private void AlertaKanbanAlmacen()
        {
            ComunicaCPROLogica.AlertaKANBAlmacen();
        }
        private void AlertaCapturaSetUp()
        {
            ComunicaCPROLogica.AlertaHORASETUP();
        }

        private void AlertaDuracionSetUp()
        {
            ComunicaCPROLogica.AlertaDURACIONSETUP();
        }
        private void AlertaRPOGlobales()
        {
            ComunicaCPROLogica.AlertaGLOBALENVIOS();
        }

        private void AlertaRPOsDetenidos()
        {
            ComunicaCPROLogica.AlertaRPOsDetenidos();
        }

        #endregion

        private void Iniciar()
        {
            string sDia = DateTime.Now.DayOfWeek.ToString();
            if (sDia == "Sunday" || sDia == "Domingo")
                return;
            else
            {
                //buscar dias festivos
                for (int i = 0; i < sFeriados.Length; i++)
                {
                    string sFecha = sFeriados[i];
                    if (DateTime.Now.ToString("dd/MM/yyyy") == sFecha)
                        return;
                }
            }


            Cursor.Current = Cursors.WaitCursor;
            CloverProEnvios();
           CargarEnviosProd();
            /*ConteoCiclico();
            
            ServicioEnvios();*/
            //CargarEnvios();

            Cursor.Current = Cursors.Arrow;
        }

        #region regCloverPro
        #region regInvCicle
        private void ConteoCiclico()
        {
            DataTable dtCon = ConfigCPROLogica.Consultar();
            string sSetup = dtCon.Rows[0]["ind_bincount"].ToString();
            string sPath = @"\\mxapp7\Interfaces\BIN";//dtCon.Rows[0]["bin_directory"].ToString();

            if (string.IsNullOrEmpty(sSetup) || sSetup == "0" || string.IsNullOrEmpty(sPath))
                return;

            //get horario
            DateTime dtTime = DateTime.Now;
            int iHora = dtTime.Hour;
            
            if (iHora > 23)
                iHora -= 23;
            int iHrStart = 6;
            int iHrEnd = 18;
             
            if (iHora < iHrStart && iHora > iHrEnd)
                return;

            if (iHora >= 0 && iHora < 6)
                dtTime = dtTime.AddDays(-1);

            string sHrReg = Convert.ToString(iHora).PadLeft(2, '0') + ":00";

            string[] sPlantas = { "0", "EMP", "TNR", "CTNR", "FUS", "INKM", "INKP" };

            for (int i = 1; i < sPlantas.Count(); i++) //plantas {EMP,TNR,CTNR,FUS,INKM,INKP}
            {
                BincontentCPROLogica bin = new BincontentCPROLogica();
                bin.Planta = sPlantas[i];
                bin.Fecha = dtTime;
                bin.Hora = sHrReg;
                if (!BincontentCPROLogica.Verificar(bin))
                {
                    string sFile = "BinContents" + i.ToString() + ".csv";
                    sFile = sPath + @"\" + sFile;
                    if (File.Exists(sFile))
                        BinContents(sPlantas[i], sHrReg, sFile);
                }
                
                string sFile2 = "ReleasedOrders" + i.ToString() + ".csv";
                sFile2 = sPath + "\\" + sFile2;
                if (File.Exists(sFile2))
                {
                    DataTable dtR = ReleasedRPO(sPlantas[i], sHrReg, sFile2);
                    if(dtR.Rows.Count > 0)
                    {
                        string sFile3 = "RegisteredPick" + i.ToString() + ".csv";
                        sFile3 = sPath + "\\" + sFile3;
                        if (File.Exists(sFile3))
                            RegisteredPick(sPlantas[i], sHrReg, sFile3, dtR);
                    }
                }
            }
        }
        private void RegisteredPick(string _asPlanta, string _asHora, string _asFile, DataTable _data)
        {
            DataTable dt = LoadFile2(_asFile);
            int iRows = dt.Rows.Count;
            
            BincontentCPROLogica bin = new BincontentCPROLogica();
            bin.Planta = _asPlanta;
            bin.Hora = _asHora;
            bin.Fecha = DateTime.Today;
            DataTable dtLines = BincontentCPROLogica.ConsultarLinea(bin);
            if(dtLines.Rows.Count > 0)
            {
                for(int b = 0; b < dtLines.Rows.Count; b++)//dtLines.Rows.Count; b++)
                {
                    string sLinea = dtLines.Rows[b][0].ToString();
                    bin.BinCode = sLinea;
                    DataTable dtBin = BincontentCPROLogica.BinContentLinea(bin);
                    //{folio,item,cantidad,contador,diferencia}
                    for (int c = 0; c < dtBin.Rows.Count; c++)
                    {
                        long lFolio = long.Parse(dtBin.Rows[c][0].ToString());
                        string sItem = dtBin.Rows[c][1].ToString();
                        double dCant = double.Parse(dtBin.Rows[c][2].ToString());
                        double dCantT = 0;

                        DataRow[] rown = dt.Select("[Bin Code]='" + sLinea+ "' AND [Item No.] = '" + sItem+"'");
                        foreach(var nrow in rown)
                        {
                            string sRPO = nrow[1].ToString();
                            bool bExis = false;
                            DataRow[] _row = _data.Select("[No.]='" + sRPO + "'");
                            if (_row.Count() > 0)
                                bExis = true;
                           
                            if (!bExis)
                                continue;
                            
                            double dCantD = 0;
                            if (!double.TryParse(nrow[5].ToString(), out dCantD)) //Quantity
                                dCantD = 0;

                            dCantT += dCantD;
                        }
                       
                        bin.Folio = lFolio;
                        bin.Contador = dCantT;
                        double dCantDif = dCant - dCantT;
                        bin.Diferencia = dCantDif; //diferencia

                        if (BincontentCPROLogica.GuardaContador(bin))
                            continue;
                    }
                }
            }
        }
        private DataTable ReleasedRPO(string _asPlanta, string _asHora, string _asFile)
        {
            DataTable dt = LoadFile(_asFile);

            for (int x = 0; x < dt.Rows.Count; x++)
            {
                
                string sRouting = dt.Rows[x][1].ToString();
                if (string.IsNullOrEmpty(sRouting))
                {
                    dt.Rows[x].Delete();
                    x--;
                    continue;
                }
            }
            return dt;
        }
        private void BinContents(string _asPlanta,string _asHora,string _asFile)
        {
            try
            {
                DataTable dt = LoadFile(_asFile);

                if (dt.Rows.Count == 0)
                    return;
                string sPta = _asPlanta;
                if (sPta == "EMP")
                    sPta = "TPACKA";

                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    string sBin = dt.Rows[x][0].ToString();
                    if (!sBin.StartsWith(sPta))
                        continue;

                    string sItem = dt.Rows[x][1].ToString();
                    string sDescrip = dt.Rows[x][2].ToString();
                    string sUM = dt.Rows[x][3].ToString();
                    double dCant = 0;
                    if (!double.TryParse(dt.Rows[x][4].ToString(), out dCant))
                        dCant = 0;
                    
                    //save data
                    BincontentCPROLogica bin = new BincontentCPROLogica();
                    bin.Hora = _asHora;
                    bin.Planta = _asPlanta;
                    bin.BinCode = sBin;
                    bin.Item = sItem;
                    bin.Descrip = sDescrip;
                    bin.UM = sUM;
                    bin.Cantidad = dCant;

                    BincontentCPROLogica.Guardar(bin);
                }
            }
            catch (Exception ex)
            {
                string sErr = ex.ToString();
                
            }
        }
        private DataTable LoadFile2(string _asFile)
        {
            int iErr = 0;
            DataTable dt = new DataTable();
            try
            {
                using (StreamReader sr = new StreamReader(_asFile))
                {
                    string[] headers = sr.ReadLine().Split(',');
                    foreach (string header in headers)
                    {
                        dt.Columns.Add(header);
                    }
                    while (!sr.EndOfStream)
                    {
                        string[] rows = sr.ReadLine().Split(',');
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i];
                            iErr = i;
                        }
                        dt.Rows.Add(dr);
                    }
                }
            }
            catch (Exception e)
            {
                string sErr = iErr.ToString() + " " + e.ToString();
            }

            return dt;
        }
        public List<string> SplitCSV(string line)
        {
            if (string.IsNullOrEmpty(line))
                throw new ArgumentException();

            List<string> result = new List<string>();

            int index = 0;
            int start = 0;
            bool inQuote = false;
            StringBuilder val = new StringBuilder();

            // parse line
            foreach (char c in line)
            {
                switch (c)
                {
                    case '"':
                        inQuote = !inQuote;
                        break;

                    case ',':
                        if (!inQuote)
                        {
                            result.Add(line.Substring(start, index - start)
                                .Replace("\"", ""));

                            start = index + 1;
                        }

                        break;
                }

                index++;
            }

            if (start < index)
            {
                result.Add(line.Substring(start, index - start).Replace("\"", ""));
            }

            return result;
        }
        private DataTable LoadFile(string _asFile)
        {
            int iErr = 0;
            DataTable dt = new DataTable();
            try
            {
                using (StreamReader sr = new StreamReader(_asFile))
                {
                    string[] headers = sr.ReadLine().Split(',');
                    foreach (string header in headers)
                    {
                        dt.Columns.Add(header);
                    }
                    while (!sr.EndOfStream)
                    {
                        
                        List<string> result = SplitCSV(sr.ReadLine());
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = result[i];
                            iErr = i;

                        }
                        dt.Rows.Add(dr);
                    }

                }
            }
            catch (Exception e)
            {
                string sErr = iErr.ToString() + " " + e.ToString();
                MessageBox.Show(sErr, Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            return dt;
        }
       
        #endregion

        private void CloverProEnvios()
        {
           // AlertaRPOsDetenidos();
            //AlertaKanban();
            //AlertaKanbanAlmacen();
            //GeneraDuracionSetUpEmp();
            //GeneraDuracionSetUp();

            //AlertaCapturaSetUp();
            //AlertaPreviaSetUp();
            //AlertaDuracionSetUp();
            //AlertaRPOGlobales();


            
            if (dgwEnvios.Rows.Count <= 0)
            {
                return;
            }

            foreach (DataGridViewRow row in dgwEnvios.Rows)
            {
                EnviarCorreoCPRO envMail = new EnviarCorreoCPRO();
                envMail.Folio = Int32.Parse(row.Cells[0].Value.ToString());
                envMail.Proceso = row.Cells[1].Value.ToString();
                envMail.Referencia = row.Cells[3].Value.ToString();
                envMail.Para = row.Cells[4].Value.ToString();
                envMail.Asunto = row.Cells[5].Value.ToString();
                envMail.Mensaje = row.Cells[7].Value.ToString();

                CorreoCPROLogica mail = new CorreoCPROLogica();
                mail.Folio = Int32.Parse(row.Cells[0].Value.ToString());
                mail.Proceso = row.Cells[1].Value.ToString();
                mail.Referencia = row.Cells[3].Value.ToString();
                
                if (EnviarCorreoCPRO.Enviar(envMail))
                {
                    mail.Estado = "E";
                }
                else
                {
                    mail.Estado = "R";
                }

                CorreoCPROLogica.Guardar(mail); //actualizar estado del correo
            }
            
        }

        private void GeneraDuracionSetUp()
        {
            //CARGAR SETUPS DEL DIA
            try
            {
                
                DataTable dtCon = ConfigCPROLogica.Consultar();
                string sSetup = dtCon.Rows[0]["setup_duration"].ToString();
                if (string.IsNullOrEmpty(sSetup) || sSetup == "0")
                    return;

                int iCant = int.Parse(dtCon.Rows[0]["setup_durturno"].ToString());
                int iMin = int.Parse(dtCon.Rows[0]["setup_durmin"].ToString());
                string sIni1t = dtCon.Rows[0]["setup_durini1t"].ToString();
                string sIni2t = dtCon.Rows[0]["setup_durini2t"].ToString();
                string sSource = dtCon.Rows[0]["setup_datasource"].ToString();
                string sSourceEmp = @"\\mxapp7\Interfaces\Setup\Backup"; //Acomulado de empaque

                MonitorSetupLogica mon = new MonitorSetupLogica();
                mon.FechaFin = DateTime.Today;
                mon.FechaIni = DateTime.Today;

                int iHoraIni = DateTime.Now.Hour;
                if(iHoraIni >= 0 && iHoraIni < 6)
                    mon.FechaIni = DateTime.Today.AddDays(-1);

                mon.IndPlanta = "0";
                mon.Planta = "";
                
                mon.IndLinea = "0";
                mon.Linea = "";

                DataTable data = MonitorSetupLogica.ListarSetup(mon);
                if(data.Rows.Count > 0)
                {
                    //busca si ya se actualizo el horario
                    string sTurno = GlobalVar.TurnoGlobal();
                    DateTime dtNow = DateTime.Now;
                    int iHora = dtNow.Hour;
                    int iHrProg = 0;

                    string sHoraIni = sIni1t; // kardex verifica ultima corrida del dia

                    string sFecha = string.Format("{0:MM/dd/yyyy HH:mm }", dtNow);
                    string sFini = string.Format("{0:MM/dd/yyyy}", dtNow);
                    if(sTurno == "1")
                        sFini += " " + sIni1t + ":00.000";
                    else
                    {
                        sFini += " " + sIni2t + ":00.000";
                        sHoraIni = sIni2t;
                    }

                    DateTime dtFini = Convert.ToDateTime(sFini); //2018-05-08 09:00:00.000

                    iHrProg = dtFini.Hour;
                    // 9am < 9am
                    if (iHora < iHrProg)
                        return;
                    else
                        iCant--;

                    bool bLoad = false;
                    while(iCant > 0) // 1
                    {
                        
                        KardexCPROLogica kar = new KardexCPROLogica();
                        kar.Proceso = "SETUP_MP";
                        kar.Fecha = DateTime.Today;

                        if ( iHora >= 0 && iHora < 6)
                            kar.Fecha = DateTime.Today.AddDays(-1);

                        kar.Hora = sHoraIni;
                        if(KardexCPROLogica.Verificar(kar))//verifica siguiente corrida
                        {
                            DateTime dtNew = dtFini.AddMinutes(iMin); // 2pm
                            iHrProg = dtNew.Hour;//2
                            if (iHora < iHrProg)
                                iCant--;
                            else
                            {
                                iCant = 0;
                                bLoad = true;
                            }
                        }
                        else
                        {
                            iCant = 0;
                            bLoad = true;//sin registro previo en el dia
                        }
                    }

                    if (!bLoad)
                        return;

                    //find setup datasource (determinar cual archivo usar)
                    sHoraIni = iHrProg.ToString().PadLeft(2, '0') + ":00";
                    iHrProg += 2;//CT zone

                    if (iHrProg < 12)
                        sFini = iHrProg.ToString() + "Am";
                    else
                    {
                        if(iHrProg > 23)
                        {
                            iHrProg -= 24;
                            sFini = iHrProg.ToString() + "Am";
                        }
                        else
                        {
                            iHrProg -= 12;
                            sFini = iHrProg.ToString() + "Pm";
                        }
                    }

                    sSource += "_" + sFini+".xlsx";
                    sSourceEmp += "_" + sFini + ".xlsx";
                    //load data soruce if exist in server
                    DataTable dtExc = new DataTable();
                    if (!File.Exists(sSource))
                        return;
                    else
                    {
                        dtExc = getFromExcel(sSource);
                        File.Copy(sSource, sSourceEmp, true);

                        File.Delete(sSource);

                        KardexCPROLogica kar = new KardexCPROLogica();
                        kar.Proceso = kar.Proceso = "SETUP_MP";
                        kar.Referencia = sSource;
                        kar.Result = dtExc.Rows.Count.ToString();
                        kar.Hora = sHoraIni;
                        kar.Usuario = "SYSCOM";
                        KardexCPROLogica.Guardar(kar);

                    }
                        
                    if (dtExc.Rows.Count == 0)
                        return;
                    
                    //setup compare
                    for (int x = 0; x< data.Rows.Count; x++)
                    {
                        long lFolio = long.Parse(data.Rows[x][1].ToString());
                        int iCons = int.Parse(data.Rows[x][2].ToString());
                        string sPta = data.Rows[x][4].ToString();
                        string sLine = data.Rows[x][5].ToString();
                        string sRPO = data.Rows[x][6].ToString();
                        string sRPOnext = data.Rows[x][7].ToString();
                        string sHora = data.Rows[x][8].ToString();
                        double dDura = double.Parse(data.Rows[x][9].ToString());
                        
                        if (sPta == "COL")
                            sPta = "MX02";
                        if (sPta == "NIC2" || sPta == "NIC3" || sPta == "MON")
                            sPta = "MX01A";
                        //if (sPta == "FUS")
                        //    sPta = "MXFSR";
                        
                        for (int c = 0; c < dtExc.Rows.Count; c++)
                        {
                            string sPtaMP = dtExc.Rows[c][0].ToString().ToUpper();
                            if (sPtaMP != sPta)
                                continue;

                            string sLineMP = dtExc.Rows[c][1].ToString();
                            if (sLineMP != sLine)
                                continue;

                            string sRpoMP = dtExc.Rows[c][2].ToString();
                            string sRpoMP2 = dtExc.Rows[c][3].ToString();
                            string sIniMP = dtExc.Rows[c][4].ToString();
                            string sFinMP = dtExc.Rows[c][5].ToString();
                            string sDura = dtExc.Rows[c][6].ToString();
                            DateTime dtHora = DateTime.Now;
                            if (!string.IsNullOrEmpty(sHora))
                                dtHora = DateTime.Parse(sHora);

                            if (sRPO == sRpoMP && sRPOnext == sRpoMP2)//VALIDA RPO X LINEA
                            {
                                //VALIDA FECHA
                                if (!string.IsNullOrEmpty(sHora))
                                {
                                    DateTime dtIni = DateTime.Parse(sIniMP);
                                    int iDays = (dtIni - dtHora).Days;
                                    if (iDays > 0)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        if (dtIni.Day != dtHora.Day)
                                        {
                                            if (dtIni.Hour >= 6)
                                                continue;
                                        }
                                    }
                                }


                                LineSetupDetLogica set = new LineSetupDetLogica();
                                set.Folio = lFolio;
                                set.Consec = iCons;

                                dtHora = DateTime.Parse(sIniMP);
                                dtHora = dtHora.AddMinutes(-120);
                                sIniMP = string.Format("{0:MM/dd/yyyy HH:mm tt}", dtHora);
                                set.IniciaMP = sIniMP;

                                dtHora = DateTime.Parse(sFinMP);
                                dtHora = dtHora.AddMinutes(-120);
                                sFinMP = string.Format("{0:MM/dd/yyyy HH:mm tt}", dtHora);
                                set.FinalMP = sFinMP;

                                set.DuraMP = Double.Parse(sDura);
                                set.RpoMP = sRpoMP2;
                                set.Usuario = "SYSCOM";

                                LineSetupDetLogica.ActualizaDuraMP(set);
                                c = dtExc.Rows.Count;
                            }
                        }
                    }
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private DataTable getFromExcel(string _asArchivo)
        {
            
            DataTable dt = new DataTable("TimeDiff");
            dt.Columns.Add("planta", typeof(string));
            dt.Columns.Add("linea", typeof(string));
            dt.Columns.Add("rpo", typeof(string));
            dt.Columns.Add("rpo_next", typeof(string));
            dt.Columns.Add("fecha", typeof(string));
            dt.Columns.Add("fecha_next", typeof(string));
            dt.Columns.Add("duracion", typeof(string));

            try
            {
                
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
                Excel.Workbook xlWorkbook = xlWorkbookS.Open(_asArchivo);

                Excel.Worksheet xlWorksheet = new Excel.Worksheet();

                string sValue = string.Empty;

                int iSheets = xlWorkbook.Sheets.Count;

                xlWorksheet = xlWorkbook.Sheets[iSheets];

                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 2; i < rowCount; i++)
                {
                    DateTime dtFecha = DateTime.Today;
                    string sRPO = string.Empty;
                    string sRPO2 = string.Empty;
                    string sPta = string.Empty;
                    string sLine = string.Empty;
                    string sFecha = string.Empty;
                    string sFecha2 = string.Empty;
                    string sMin = string.Empty;

                    sValue = string.Empty;

                    if (xlRange.Cells[i, 8].Value2 == null)
                        continue;

                    if (xlRange.Cells[i, 1].Value2 != null)
                        sValue = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());


                    if (sValue == "TimeDiffBetweenRPOs" || sValue == "Facility Name")
                    {
                        sValue = string.Empty;
                        continue;
                    }

                    if (string.IsNullOrEmpty(sValue))
                        continue;

                    sPta = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());
                    sLine = Convert.ToString(xlRange.Cells[i, 2].Value2.ToString());
                    sPta = sPta.ToUpper();
                    //if (sPta != "MX01A" && sPta != "MX02" && sPta != "MX06" && sPta != "MXFSR")
                    if (sPta != "MX01A" && sPta != "MX02")
                        continue;

                    int iLine = 0;
                    if(int.TryParse(sLine, out iLine))
                    {
                        if (iLine > 55)
                            continue;
                    }

                    if (sPta == "MX02")
                        sLine = "C" + sLine.PadLeft(2,'0');

                    sRPO = Convert.ToString(xlRange.Cells[i, 3].Value2.ToString());
                    sRPO = sRPO.TrimStart().TrimEnd().ToUpper();
                    sRPO2 = Convert.ToString(xlRange.Cells[i, 4].Value2.ToString());
                    sRPO2 = sRPO2.TrimStart().TrimEnd().ToUpper();
                    if (xlRange.Cells[i, 5].Value2 != null)
                        sFecha = Convert.ToString(xlRange.Cells[i, 5].Value.ToString());
                    else
                        sFecha = Convert.ToString(xlRange.Cells[i, 6].Value.ToString());
                    sFecha2 = Convert.ToString(xlRange.Cells[i, 7].Value.ToString());
                    sMin = Convert.ToString(xlRange.Cells[i, 8].Value2.ToString());

                    dt.Rows.Add(sPta, sLine, sRPO, sRPO2, sFecha, sFecha2, sMin);

                }

                xlApp.DisplayAlerts = false;
                xlWorkbook.Close();
                xlApp.DisplayAlerts = true;
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                ex.ToString();
                

            }

            return dt;
        }

        private void CargarEnviosProd()
        {
            DataTable dtEnv = new DataTable();
            dtEnv = ComunicaCPROLogica.EnviosPendientes();
            dgwEnvios.DataSource = dtEnv;
            CargarColumnas();

        }

        #region regSetupEmp
        private DataTable getFromExcelEmp(string _asArchivo)
        {

            DataTable dt = new DataTable("TimeDiff");
            dt.Columns.Add("planta", typeof(string));
            dt.Columns.Add("linea", typeof(string));
            dt.Columns.Add("rpo", typeof(string));
            dt.Columns.Add("rpo_next", typeof(string));
            dt.Columns.Add("fecha", typeof(string));
            dt.Columns.Add("fecha_next", typeof(string));
            dt.Columns.Add("duracion", typeof(string));

            try
            {

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
                Excel.Workbook xlWorkbook = xlWorkbookS.Open(_asArchivo);

                Excel.Worksheet xlWorksheet = new Excel.Worksheet();

                string sValue = string.Empty;

                int iSheets = xlWorkbook.Sheets.Count;

                xlWorksheet = xlWorkbook.Sheets[iSheets];

                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 2; i < rowCount; i++)
                {
                    DateTime dtFecha = DateTime.Today;
                    string sRPO = string.Empty;
                    string sRPO2 = string.Empty;
                    string sPta = string.Empty;
                    string sLine = string.Empty;
                    string sFecha = string.Empty;
                    string sFecha2 = string.Empty;
                    string sMin = string.Empty;

                    sValue = string.Empty;

                    if (xlRange.Cells[i, 8].Value2 == null)
                        continue;

                    if (xlRange.Cells[i, 1].Value2 != null)
                        sValue = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());


                    if (sValue == "TimeDiffBetweenRPOs" || sValue == "Facility Name")
                    {
                        sValue = string.Empty;
                        continue;
                    }

                    if (string.IsNullOrEmpty(sValue))
                        continue;

                    sPta = Convert.ToString(xlRange.Cells[i, 1].Value2.ToString());
                    sLine = Convert.ToString(xlRange.Cells[i, 2].Value2.ToString());
                    sPta = sPta.ToUpper();
                    if (sPta != "MX06")
                        continue;

                    sRPO = Convert.ToString(xlRange.Cells[i, 3].Value2.ToString());
                    sRPO = sRPO.TrimStart().TrimEnd().ToUpper();
                    sRPO2 = Convert.ToString(xlRange.Cells[i, 4].Value2.ToString());
                    sRPO2 = sRPO2.TrimStart().TrimEnd().ToUpper();
                    if (xlRange.Cells[i, 5].Value2 != null)
                        sFecha = Convert.ToString(xlRange.Cells[i, 5].Value.ToString());
                    else
                        sFecha = Convert.ToString(xlRange.Cells[i, 6].Value.ToString());
                    sFecha2 = Convert.ToString(xlRange.Cells[i, 7].Value.ToString());
                    sMin = Convert.ToString(xlRange.Cells[i, 8].Value2.ToString());

                    dt.Rows.Add(sPta, sLine, sRPO, sRPO2, sFecha, sFecha2, sMin);

                }

                xlApp.DisplayAlerts = false;
                xlWorkbook.Close();
                xlApp.DisplayAlerts = true;
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                ex.ToString();


            }

            return dt;
        }

        private void GeneraDuracionSetUpEmp()
        {
            //CARGAR SETUPS DEL DIA
            try
            {

                DataTable dtCon = ConfigCPROLogica.Consultar();
                string sSetup = dtCon.Rows[0]["setup_duration"].ToString();
                if (string.IsNullOrEmpty(sSetup) || sSetup == "0")
                    return;
                //PARAMETROS EMPAQUE
                string sIni1t = "06:00";//dtCon.Rows[0]["setup_durini1t"].ToString();
                string sSource = @"\\mxapp7\Interfaces\Setup\Backup"; //Acomulado de empaque
                string sFile = @"\TimeDiffBetweenRPOs_8Pm.xlsx";

                string sTurno = GlobalVar.TurnoGlobal();

                DateTime dtNow = DateTime.Now;
                int iHora = dtNow.Hour;
                int iHrProg = 0;

                string sHoraIni = sIni1t; // kardex verifica ultima corrida del dia

                string sFecha = string.Format("{0:MM/dd/yyyy HH:mm }", dtNow);

                //CONVERTIR HORARIO PROGRAMADO EN DATETIME
                string sFini = string.Format("{0:MM/dd/yyyy}", dtNow);
                sFini += " " + sIni1t + ":00.000";
                DateTime dtFini = Convert.ToDateTime(sFini); //2018-05-08 09:00:00.000

                iHrProg = dtFini.Hour; //6AM
                // GENERAR EL ARCHIVO A LAS 6AM
                if (iHora != iHrProg)
                    return;

                KardexCPROLogica kar = new KardexCPROLogica();
                kar.Proceso = "SETUP_MPE";
                kar.Fecha = DateTime.Today;
                kar.Hora = sIni1t;
                if (!KardexCPROLogica.Verificar(kar))
                {
                    //busca si ya se GENERO EL ARCHIVO DEL DIA
                    sSource += sFile;

                    //load data soruce if exist in server
                    DataTable dtExc = new DataTable();
                    if (!File.Exists(sSource))
                        return;
                    else
                    {
                        dtExc = getFromExcelEmp(sSource);

                        kar.Proceso = kar.Proceso = "SETUP_MPE";
                        kar.Referencia = sSource;
                        kar.Result = dtExc.Rows.Count.ToString();
                        kar.Hora = sHoraIni;
                        kar.Usuario = "SYSCOM";
                        KardexCPROLogica.Guardar(kar);

                    }

                    if (dtExc.Rows.Count == 0)
                        return;

                    ExportarTexto(sSource, sFile, dtExc);
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }
        private string GetFecha(int _aiTipo,int _aiDia)
        {
            string sValor = string.Empty;
            DateTime dtTime = DateTime.Today;

            if (_aiDia != 0)
                dtTime = DateTime.Today.AddDays(_aiDia);

            if (_aiTipo == 1)
                sValor = Convert.ToString(dtTime);

            if (_aiTipo == 2)
                sValor = Convert.ToString(dtTime.Month);

            if (_aiTipo == 3)
            {
                sValor = Convert.ToString(dtTime.Month);
                if (sValor == "1")
                    sValor = "ENERO";
                if (sValor == "2")
                    sValor = "FEBRERO";
                if (sValor == "3")
                    sValor = "MARZO";
                if (sValor == "4")
                    sValor = "ABRIL";
                if (sValor == "5")
                    sValor = "MAYO";
                if (sValor == "6")
                    sValor = "JUNIO";
                if (sValor == "7")
                    sValor = "JULIO";
                if (sValor == "8")
                    sValor = "AGOSTO";
                if (sValor == "9")
                    sValor = "SEPTIEMBRE";
                if (sValor == "10")
                    sValor = "OCTUBRE";
                if (sValor == "11")
                    sValor = "NOVIEMBRE";
                if (sValor == "12")
                    sValor = "DICIEMBRE";
            }

            return sValor;
        }
        private void ExportarTexto(string _asDire, string _asFile, DataTable _aDt)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string sDia = GetFecha(1,-1);//CALCULAR ULT DIA HABIL (VIERNES-LUNES)
                string sMes = GetFecha(2,-1);
                string sMesDesc = GetFecha(3,-1);
                string sTurno = GlobalVar.TurnoGlobal();

                _asDire = @"\\mxapp7\Interfaces\Setup\Packaging";
                _asDire += "\\" + sMesDesc;
                _asFile = "LINESETUP_PACKING";

                bool bExists = Directory.Exists(_asDire);
                if (!bExists)
                    Directory.CreateDirectory(_asDire);

                _asFile += "_T" + sTurno + "_" + sDia.PadLeft(2, '0') + "-" + sMes.PadLeft(2, '0');
                
                string sFile = _asDire + "\\" + _asFile + ".csv";
                using (var stream = File.CreateText(sFile))
                {
                    int iCantT = 0;
                    int iCantL = 0;
                    int iCant1 = 0;
                    int iCant2 = 0;
                    int iCant3 = 0;
                    string sLineAnt = string.Empty;
                    double dMinsL = 0;
                    double dMins1 = 0;
                    double dMins2 = 0;
                    double dMins3 = 0;
                    for (int x = 0; x < _aDt.Rows.Count; x++)
                    {
                        iCantT++;
                        string sPta = _aDt.Rows[x][0].ToString().ToUpper();
                        string sLine = _aDt.Rows[x][1].ToString();
                        string sRpo = _aDt.Rows[x][2].ToString();
                        string sDura = _aDt.Rows[x][6].ToString();

                        if (string.IsNullOrEmpty(sLineAnt))
                            sLineAnt = sLine;

                        if (sLineAnt == sLine)
                            iCantL++;
                        else
                            iCantL = 0;

                        //genera archivo excel son formato hperez
                        

                        string sRow = string.Format("{0},{1},{2},{3}",sLine, iCant1, iCant2, iCant3);

                        stream.WriteLine(sRow);

                        sLineAnt = sLine;
                    }

                    stream.Close();

                    if (File.Exists(sFile))
                    {

                        KardexCPROLogica kar = new KardexCPROLogica();
                        kar.Proceso = "SETUP-MPE";
                        kar.Referencia = _asFile;
                        kar.Result = "OK";
                        kar.Hora = "06:00";

                        KardexCPROLogica.Guardar(kar);
                    }
                }

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                string sEx = ex.ToString();
                throw;
            }

        }
        #endregion

        #endregion

        private void ServicioEnvios()
        {
            AlertaMaqCosto();

            AlertaProductividad();

            AlertaTOP10CtoServ();

            AlertaTOP10Servicios();

            AlertaOSFECOM();

            AlertaFechaPromesa();

            AlertaExistencia();

            //Listar correos pendientes de enviar
            CargarEnvios();

            if (dgwEnvios.Rows.Count <= 0)
                return;

            foreach (DataGridViewRow row in dgwEnvios.Rows)
            {
                EnviarCorreo envMail = new EnviarCorreo();
                envMail.Folio = Int32.Parse(row.Cells[0].Value.ToString());
                envMail.Proceso = row.Cells[1].Value.ToString();
                envMail.Referencia = row.Cells[3].Value.ToString();
                envMail.Para = row.Cells[4].Value.ToString();
                envMail.Asunto = row.Cells[5].Value.ToString();
                envMail.Mensaje = row.Cells[7].Value.ToString();

                CorreoLogica mail = new CorreoLogica();
                mail.Folio = Int32.Parse(row.Cells[0].Value.ToString());
                mail.Proceso = row.Cells[1].Value.ToString();
                mail.Referencia = row.Cells[3].Value.ToString();

                if (EnviarCorreo.Enviar(envMail))
                {
                    mail.Estado = "E";
                }
                else
                {
                    mail.Estado = "R";
                }

                CorreoLogica.Guardar(mail); //actualizar estado del correo
            }
        }

        private void CargarEnvios()
        {
            DataTable dtEnv = new DataTable();
            dtEnv = ComunicaLogica.EnviosPendientes();
            dgwEnvios.DataSource = dtEnv;
            CargarColumnas();

        }

        private void CargarEnviados()
        {
            DataTable dtEnv2 = new DataTable();
            dtEnv2 = ComunicaLogica.Enviados();
            dgwEnviados.DataSource = dtEnv2;
            CargarColumnas2();

        }

        private void CargarColumnas2()
        {
            if (dgwEnviados.Rows.Count == 0)
            {
                DataTable dtNew = new DataTable("Envios");
                dtNew.Columns.Add("f_id", typeof(DateTime));
                dtNew.Columns.Add("Folio", typeof(long));
                dtNew.Columns.Add("Proceso", typeof(string));
                dtNew.Columns.Add("Destinatario", typeof(string));
                dtNew.Columns.Add("Asunto", typeof(string));
            }

            dgwEnviados.Columns[0].Width = ColumnWith(dgwEnvios, 15);
            dgwEnviados.Columns[1].Width = ColumnWith(dgwEnvios, 5);
            dgwEnviados.Columns[2].Width = ColumnWith(dgwEnvios, 15);
            dgwEnviados.Columns[3].Width = ColumnWith(dgwEnvios, 15);
            dgwEnviados.Columns[4].Width = ColumnWith(dgwEnvios, 48);
        }
        private void CargarColumnas()
        {
            if (dgwEnvios.Rows.Count == 0)
            {
                DataTable dtNew = new DataTable("Envios");
                dtNew.Columns.Add("Folio", typeof(long));
                dtNew.Columns.Add("proceso", typeof(string));
                dtNew.Columns.Add("NombreProceso", typeof(string));
                dtNew.Columns.Add("Referencia", typeof(string));
                dtNew.Columns.Add("Destinatario", typeof(string));
                dtNew.Columns.Add("Asunto", typeof(string));
                dtNew.Columns.Add("Estado", typeof(string));
                dtNew.Columns.Add("mensaje", typeof(string));
                dtNew.Columns.Add("u_id", typeof(string));
            }
            
            dgwEnvios.Columns[1].Visible = false;
            dgwEnvios.Columns[7].Visible = false;
            dgwEnvios.Columns[8].Visible = false;

            dgwEnvios.Columns[0].Width = ColumnWith(dgwEnvios, 5);
            dgwEnvios.Columns[2].Width = ColumnWith(dgwEnvios, 15);
            dgwEnvios.Columns[3].Width = ColumnWith(dgwEnvios, 10);
            dgwEnvios.Columns[4].Width = ColumnWith(dgwEnvios, 30);
            dgwEnvios.Columns[5].Width = ColumnWith(dgwEnvios, 32);
            dgwEnvios.Columns[6].Width = ColumnWith(dgwEnvios, 8);
        }

        #region regControles

        private void btnRun_Click(object sender, EventArgs e)
        {
            Iniciar();
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            
        }

        private void dgwEnvios_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            int iRow = e.RowIndex;
            if ((iRow % 2) == 0)
                e.CellStyle.BackColor = Color.WhiteSmoke;
            else
                e.CellStyle.BackColor = Color.White;
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (dgwEnvios.SelectedRows.Count == 0)
                return;
            
            if (!string.IsNullOrEmpty(dgwEnvios.SelectedCells[0].Value.ToString()))
            {
                ComunicaLogica com = new ComunicaLogica();
                com.Folio = long.Parse(dgwEnvios.SelectedCells[0].Value.ToString());
                ComunicaLogica.Eliminar(com);
            }
            dgwEnvios.Rows.Remove(dgwEnvios.CurrentRow);
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage == tabPage1 )
                CargarEnvios();
            if (e.TabPage == tabPage2)
                CargarEnviados();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dgwEnvios.SelectedRows.Count == 0)
                return;

            DialogResult Result = MessageBox.Show("Se borrara todo el registro de los correos enviados. Desea Continuar?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (Result == DialogResult.Yes)
            {
                Cursor.Current = Cursors.WaitCursor;

                ComunicaLogica.EliminarEnviados();

                Cursor.Current = Cursors.Arrow;
                CargarEnviados();

            }
        }

        private void dgwEnvios_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            
        }

        #endregion
        private void timer1_Tick(object sender, EventArgs e)
        {
            Iniciar();
        }
    }
}
