using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Data;

namespace Logica
{
    public class EnviarCorreo
    {
        public long Folio { get; set; }
        public string Proceso { get; set; }
        public string Referencia { get; set; }
        public string Para { get; set; }
        public string Asunto { get; set; }
        public string Mensaje { get; set; }

        public static bool Enviar(EnviarCorreo enviar)
        {

            DataTable dtMail = new DataTable();
            dtMail = ConfigCorreoLogica.ConsultarSMTP();
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


                    message.From = new MailAddress(dtMail.Rows[0]["correo"].ToString());

                    bool bAlerta = false;
                    long lRefer;
                    if (!long.TryParse(enviar.Referencia, out lRefer))
                        bAlerta = true;

                    if (enviar.Proceso == "A005" || enviar.Proceso == "H010" || bAlerta)
                        message.To.Add(enviar.Para);
                    else
                    {
                        DataTable dtTo = new DataTable();
                        //ORIGEN FOLIO NUMERICO [ORDEN SERVICIO]
                        if (enviar.Proceso == "H037") //Mensajes de Solicitudes
                            enviar.Proceso = "C060";

                        dtTo = ConfigDestLogica.Consultar(enviar.Proceso);
                        for (int i = 0; i <= dtTo.Rows.Count - 1; i++)
                        {
                            if (dtTo.Rows[i]["tipo"].ToString() == "T")
                                message.To.Add(dtTo.Rows[i]["correo"].ToString());
                            else
                                message.CC.Add(dtTo.Rows[i]["correo"].ToString());
                        }
                    }

                    message.Subject = enviar.Asunto;
                    string sMensaje = "";
                    if (enviar.Proceso == "A005" || enviar.Proceso == "C060" || enviar.Proceso == "H037" || bAlerta) //Mensajes en HTML
                    {
                         DataTable dtBody = CorreoLogica.BodyMail(enviar.Folio);
                        for(int i = 0; i <= dtBody.Rows.Count -1; i++)
                        {
                            string sBody = dtBody.Rows[i]["body"].ToString();
                            sMensaje += sBody;
                        }
                        message.Body = sMensaje;
                    }
                    else
                        message.Body = enviar.Mensaje;

                    if (dtMail.Rows[0]["ind_html"].ToString() == "1")
                        message.IsBodyHtml = true;
                    else
                        message.IsBodyHtml = false;

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
