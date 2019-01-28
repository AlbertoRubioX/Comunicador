using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Data;

namespace Logica
{
    public class EnviarCorreoCPRO
    {
        public long Folio { get; set; }
        public string Proceso { get; set; }
        public string Referencia { get; set; }
        public string Para { get; set; }
        public string Asunto { get; set; }
        public string Mensaje { get; set; }
        public static bool Enviar(EnviarCorreoCPRO enviar)
        {

            DataTable dtMail = new DataTable();
            dtMail = ConfigCPROLogica.Consultar();
            if (dtMail.Rows.Count > 0)
            {
                try
                {
                    MailMessage message = new MailMessage();
                    SmtpClient smtp = new SmtpClient();

                    message.IsBodyHtml = true;
                    smtp.Host = dtMail.Rows[0]["servidor"].ToString();
                    smtp.Port = Int32.Parse(dtMail.Rows[0]["puerto"].ToString());
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential(dtMail.Rows[0]["usuario"].ToString(), dtMail.Rows[0]["password"].ToString());

                    if (dtMail.Rows[0]["ind_ssl"].ToString() == "1")
                        smtp.EnableSsl = true;
                    else
                        smtp.EnableSsl = false;

                    message.From = new MailAddress(dtMail.Rows[0]["correo_sal"].ToString());

                    bool bAlerta = false;
                    long lRefer;
                    if (!long.TryParse(enviar.Referencia, out lRefer))
                        bAlerta = true;
                                        
                    message.Subject = enviar.Asunto;
                    string sMensaje = "";
                    if (enviar.Proceso == "PRO050" || enviar.Proceso == "PLA020" || enviar.Proceso == "EMP040" || enviar.Proceso == "KANBAN" || enviar.Proceso == "KANBALM" || enviar.Proceso == "EMP050" ) //Mensajes en HTML
                    {
                        DataTable dtBody = CorreoCPROLogica.BodyMail(enviar.Folio);
                        for(int i = 0; i <= dtBody.Rows.Count -1; i++)
                        {
                            string sBody = dtBody.Rows[i]["body"].ToString();
                            sMensaje += sBody;
                        }
                        message.Body = sMensaje;
                    }
                    else
                        message.Body = enviar.Mensaje;

                    DataTable dtTo = CorreoCPROLogica.ConsultaTo(enviar.Folio);
                    for (int i = 0; i <= dtTo.Rows.Count - 1; i++)
                    {
                        if (dtTo.Rows[i]["tipo"].ToString() == "T")
                            message.To.Add(dtTo.Rows[i]["correo"].ToString());
                        else
                            message.CC.Add(dtTo.Rows[i]["correo"].ToString());
                    }

                    if (dtMail.Rows[0]["ind_html"].ToString() == "1")
                        message.IsBodyHtml = true;
                    else
                        message.IsBodyHtml = false;

                    if (enviar.Proceso == "KANBAN")
                    {
                        string sPath = dtMail.Rows[0]["kanban_direc"].ToString();
                        string sFile = sPath + "\\KANBAN_RESUMEN_" + enviar.Referencia + ".xlsx";
                        if (File.Exists(sFile))
                        {
                            Attachment myAttachment = new Attachment(sFile);
                            message.Attachments.Add(myAttachment);
                        }
                    }


                    smtp.Send(message);
                    message.Dispose();

                    return true;


                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return false;
        }
    }
}
