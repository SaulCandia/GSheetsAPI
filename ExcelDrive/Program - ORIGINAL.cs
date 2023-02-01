using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Timers;

namespace ExcelDrive
{
    class Program
    {
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        //static string ApplicationName = "Google Sheets w/API .NET";

        static void Main(string[] args)
        {
            //Toma de argumentos
            int argumentoDias = 0;


            if (args != null || args.Length == 0)
            {
                try
                {
                    argumentoDias = int.Parse(args[0].ToString());
                }
                catch (Exception)
                {
                    Console.WriteLine("Parámetro debe ser un entero (Días de anticipación)");
                    Console.WriteLine("El programa se cerrará en dos (2) segundos para controlar esta excepción. Debe iniciarla de nuevo.");
                    System.Threading.Thread.Sleep(2000);
                    Environment.Exit(0);
                }
            }
            else
            {
                //Si no hay argumentos, el default es de tres días antes del vencimiento del proyecto
                argumentoDias = 3;
            }

            //Introducción / Presentación
            Console.WriteLine("SERVICIO FANTASMA DE SEGUIMIENTO DE TRACKER");
            Console.WriteLine("====================================================================================");
            Console.WriteLine("Saúl Candia - Área Transformación Digital, Grupo Leonera");
            //String spreadsheetId = "1i3e7TnxZQZ8Ta6nwOxK7LRkR46wz6P-V-KTB8lqIp4k"; /*ID copia Tracker*/
            String spreadsheetId = "1WW216OeWLhhHkCo4_4o3NxBJF-ErZcYzfdnqOPZgO6k"; /*ID Tracker original*/

            try
            {
                //Conecta a API

                Console.WriteLine("Conectando con API Google Sheets");
                string[] scopes = new string[] { SheetsService.Scope.Drive };
                String serviceAccountEmail = "gsheets-documental@plucky-zodiac-346217.iam.gserviceaccount.com";

                //Crea instancia de conexión
                var initializer = new ServiceAccountCredential.Initializer(serviceAccountEmail)
                {
                    Scopes = scopes,
                    User = "swdocumentalcg@leonera.cl"
                };

                //Consigue clave privada desde servicios JSON
                var credential = new ServiceAccountCredential(initializer.FromPrivateKey("-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQChOkyhxKWg3QCs\nsYiScSVhq+v21DKSWQeMFQkgizEfTNNO4VSPu61nXqhnSNMr3r74ca1rGwTEslAy\no23/7Ng+IOUpfjVXZ4hydoZdEdsRM3X5az97VYOOR+tJEtBPfBa0WQIobJHGp7gF\nRnZX3jbEbVAkG3AdU8TQUluURbkvTAOsMqt5KWtrGpar5/GLlZ3HV3WwP3fO+lpd\nZGApvAsYEij8dU5h6D0A7ccEYj8+qMXIDT3gupnpitMcm+VUu+NlwxeevZlO+WJU\nIVH4fHWNzG15R5EURukEpbo+0lgDFZ+TLS+tYi23zgjOJp+nr2alHITZhBb9/Ig/\nlgG0bfMvAgMBAAECggEALQLiLw1/8hORIyVjTAMDnSuKsnvWbI4ncb/Trv69JZBk\ns/Jrkb8jL6c5G7C0p9xFc4YFFNBTufhQNHr09EyyqFG1uKpQCQlSCia151jbUIeN\n6aa779pVYo0IjnuOpYouqoXo+NEqt4vOb8aWtnxGzPr5s0Lnv4BKA6DiiVgX1bCR\nDJkdO+x84mXY8cf7Q+gyBhD1NDRfgu+iECsLS11oquKNE607ZHTi3ML9T7xxOpcw\nwNuFFavTqnJ74G/aUyBLdkvmLUq39EQtXVeK8G907FXHXSoglywUcnFW++HhqU/k\neQrcJDfJRRuxopGrcrvVk3kiMYBqckbVVO5rxQN5SQKBgQDUzfz0vZc2zcQ7bpqV\npag+zU9g1fuPLbdlv6ZB9U9RmAX5duIBBxm2uixEkSvg4j2bzKCdT2I1SC5D02nk\n6is9RegQlBVK9GYgCf+sbccAaAFRE0KFEWAm+3xcg4oSy6CNZIQ+Ku0Uqtu72kYY\nEHqI+3rjlI5DiZ09v5Zsfg2/BwKBgQDB9DWlbKP8FuwgU04kUUeKsn31SPIm441r\nbEJjtcaLTy9zzikG4fTEOQg1dtxVCsOPHSI0t3laWXveY62xJz/C6w17bByGa0Bl\n4mJI2pM+JiPvTAJZlZo8RHKWAqDLqR5W1WxD/zHhtSt14b8aZoEYdzPB5zpfTm/K\n8pD2F5j4mQKBgDlILWQPuKl23/CDiDbp/YzSJSDS2MEktC4+VVmB19UFz+3js1hF\negV2vb3DOgVxwNW0UjOmD7B5+oIlYWbOJc97hskXo1emy+qp5lmavyt704boYUqC\nb9hub35TphIDH/ePbA1z7pdWmolJav7FSMagsuaZsWW6oEnjzXDsyXR3AoGAcShl\n1CnaUs2c3g88W/v/3W/eBSmV/hJtA+uZoEsBl22PpeT2Esnp4EHWBDtguU0aY3j/\n5/nTl1714f4N7HmVvccdipC847/XRpoZ9Z9woKXn+UlDZbjez6Kvp83Iuonk5YyH\nKfTNyX3F6XTX5jM/xmJllA+wAsLkfmefI7UIzqECgYBqi8TYxt+G8NUzUDNxzsJA\nGqqcPtcoYICpYKloLnh5PAgTXaNs0hMQU8n4Zi4ZYfSVprkTRAiW1xHdnps5QhN9\n2lLKIMeMoEdWh/j293V629UahYZtMcnbfDk3fRJukaDtmG/XPg7e29MfJCE2sJ7S\nBZ4+a+l71o8IbxiTnO/BLg==\n-----END PRIVATE KEY-----\n"));

                //Crea una instancia con el servicio API
                Google.Apis.Sheets.v4.SheetsService service = new Google.Apis.Sheets.v4.SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    //ApplicationName = "GoogleDriveRestAPI-v3",
                    ApplicationName = "DriveAPI",
                });
                service.HttpClient.Timeout = TimeSpan.FromMinutes(100);

                //String spreadsheetId = "1i4xPLbs_92bOlKdCCCTI54htzMYCdOkeCBOY6ZdRfu0";
                //String range = "Hoja 1!A:B";

                String range = "Tracker!O:O";

                SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

                var response = request.Execute();

                IList<IList<Object>> values = response.Values;

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nConectado exitosamente a la API\n");
                Console.WriteLine("ID Documento Tracker: " + spreadsheetId);
                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nLeyendo estado de proyectos... espere por favor...");


                int numFila = 1;
                int numFilaCorreo = 1;
                int atrasadoPosicionY = 0;
                int porVencerPosicionY = 0;
                int limite = 6000;
                int contadorEnCurso = 0;
                int[] posFila = new int[limite];
                string[] nombreClienteInterno = new string[limite];
                string[] nombreReponsable = new string[limite];
                string[] nombreProyecto = new string[limite];
                string[] correoResponsable = new string[limite];
                DateTime[] fechaCompromiso = new DateTime[limite];
                var hoy = DateTime.Today;

                //PARTE 1: REVISA EL ESTADO DE TAREAS Y DETERMINA POSICIÓN X EN LAS MISMAS PARA LAS TAREAS ATRASADAS O QUE VAN A VENCER

                if (values != null && values.Count > 0)
                {
                    foreach (var rowEstado in values)
                    {
                        if (rowEstado[0].ToString() == "Atrasado")
                        {
                            posFila[atrasadoPosicionY] = numFila;
                            //Console.WriteLine(posFila[posicionY]);
                            atrasadoPosicionY = atrasadoPosicionY + 1;
                        }
                        else if (rowEstado[0].ToString() == "En Curso")
                        {
                            posFila[porVencerPosicionY] = numFila;
                            //Console.WriteLine(posFila[posicionY]);
                            porVencerPosicionY = porVencerPosicionY + 1;
                        }
                        numFila = numFila + 1;
                    }
                }
                else
                {
                    Console.WriteLine("No hay información para mostrar.");
                }

                //PARTE 2: CHEQUEA EL PROYECTO POR VENCER Y SU RESPONSABLE

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nRevisando proyectos por vencer y sus responsables... espere por favor... \n");


                range = "Tracker!G:G";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                response = request.Execute();
                values = response.Values;

                if (values != null && values.Count > 0)
                {
                    //Busca aplicaciones POR VENCER
                    for (int i = 0; i < porVencerPosicionY; i++)
                    {
                        numFilaCorreo = 1;
                        range = "Tracker!G" + posFila[i].ToString() + ":L" + posFila[i].ToString();
                        request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                        response = request.Execute();
                        values = response.Values;

                        foreach (var rowProyecto in values)
                        {
                            Boolean result = false;
                            nombreProyecto[i] = rowProyecto[0].ToString();
                            fechaCompromiso[i] = Convert.ToDateTime(rowProyecto[2]);
                            nombreClienteInterno[i] = rowProyecto[4].ToString();
                            nombreReponsable[i] = rowProyecto[5].ToString();

                            //Cálculo de fechas
                            var fechaDeCompromiso = fechaCompromiso[i];
                            var diferenciaDias = fechaDeCompromiso - hoy;

                            //Console.WriteLine("\n"+fechaDeCompromiso+"\n");
                            //Console.WriteLine("\n" + diferenciaDias.Days + "\n");

                            if(diferenciaDias.Days <= argumentoDias)
                            {
                                range = "DataCorreo!A:A";
                                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                                response = request.Execute();
                                values = response.Values;

                                foreach (var rowNombreResponsable in values)
                                {
                                    if (rowNombreResponsable[0].ToString() == nombreReponsable[i])
                                    {
                                        range = "DataCorreo!B" + numFilaCorreo.ToString() + ":B" + numFilaCorreo.ToString();
                                        request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                                        response = request.Execute();
                                        values = response.Values;

                                        foreach (var rowCorreo in values)
                                        {
                                            correoResponsable[i] = rowCorreo[0].ToString();

                                            Console.WriteLine("====================================================================================");
                                            Console.WriteLine("\nSe notificará a " + correoResponsable[i].ToString() + " por proyecto en curso");
                                            Console.WriteLine("'" + nombreProyecto[i].ToString() + "'");
                                            Console.WriteLine("El cual vence el "+ Convert.ToDateTime(rowProyecto[2]).ToString("dd-MMMM-yyyy") + " (en " + diferenciaDias.Days + " días).\n");
                                            Console.WriteLine("====================================================================================");
                                            Console.WriteLine("Por favor, espere un momento mientras se envía el correo...\n");

                                            result = SendMail(nombreProyecto[i].ToString()
                                                                        , Convert.ToDateTime(fechaCompromiso[i].ToString())
                                                                        , nombreReponsable[i].ToString()
                                                                        , nombreClienteInterno[i].ToString()
                                                                        , correoResponsable[i].ToString()
                                                                        , 1); //Próximo a vencer

                                            if (result == true)
                                            {
                                                Console.WriteLine("\nEl correo fue enviado exitosamente. :-)");
                                                Console.WriteLine("====================================================================================");
                                            }
                                            else
                                            {
                                                Console.WriteLine("\nEl correo no fue enviado.\nProbablemente se deba a que no hay correo definido para el responsable o porque el mismo no existe.");
                                            }
                                        }
                                    }
                                    numFilaCorreo = numFilaCorreo + 1;
                                }
                            }
                            else
                            {
                                contadorEnCurso = contadorEnCurso + 1;
                            }
                        }
                    }

                    //Busca aplicaciones VENCIDAS
                    for (int i = 0; i < atrasadoPosicionY; i++)
                    {
                        numFilaCorreo = 1;
                        range = "Tracker!G" + posFila[i].ToString() + ":L" + posFila[i].ToString();
                        request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                        response = request.Execute();
                        values = response.Values;

                        foreach (var rowProyecto in values)
                        {
                            Boolean result = false;
                            nombreProyecto[i] = rowProyecto[0].ToString();
                            fechaCompromiso[i] = Convert.ToDateTime(rowProyecto[2]);
                            nombreClienteInterno[i] = rowProyecto[4].ToString();
                            nombreReponsable[i] = rowProyecto[5].ToString();

                            //Cálculo de fechas
                            var fechaDeCompromiso = fechaCompromiso[i];
                            var diferenciaDias = fechaDeCompromiso - hoy;

                            //Console.WriteLine("\n"+fechaDeCompromiso+"\n");
                            //Console.WriteLine("\n" + diferenciaDias.Days + "\n");

                            if (diferenciaDias.Days <= argumentoDias)
                            {
                                range = "DATA!A:A";
                                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                                response = request.Execute();
                                values = response.Values;

                                foreach (var rowNombreResponsable in values)
                                {
                                    if (rowNombreResponsable[0].ToString() == nombreReponsable[i])
                                    {
                                        range = "DATA!B" + numFilaCorreo.ToString() + ":B" + numFilaCorreo.ToString();
                                        request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                                        response = request.Execute();
                                        values = response.Values;

                                        foreach (var rowCorreo in values)
                                        {
                                            correoResponsable[i] = rowCorreo[0].ToString();

                                            Console.WriteLine("====================================================================================");
                                            Console.WriteLine("\nSe notificará a " + correoResponsable[i].ToString() + " por proyecto vencido");
                                            Console.WriteLine("'" + nombreProyecto[i].ToString() + "'");
                                            Console.WriteLine("El cual venció el " + Convert.ToDateTime(rowProyecto[2]).ToString("dd-MMMM-yyyy") + " (hace " + diferenciaDias.Days + " días).\n");
                                            Console.WriteLine("====================================================================================");
                                            Console.WriteLine("Por favor, espere un momento mientras se envía el correo...\n");

                                            result = SendMail(nombreProyecto[i].ToString()
                                                                        , Convert.ToDateTime(fechaCompromiso[i].ToString())
                                                                        , nombreReponsable[i].ToString()
                                                                        , nombreClienteInterno[i].ToString()
                                                                        , correoResponsable[i].ToString()
                                                                        , 0); //Atrasado

                                            if (result == true)
                                            {
                                                Console.WriteLine("\nEl correo fue enviado exitosamente. :-)");
                                                Console.WriteLine("====================================================================================");
                                            }
                                            else
                                            {
                                                Console.WriteLine("\nEl correo no fue enviado.\nProbablemente se deba a que no hay correo definido para el responsable o porque el mismo no existe.");
                                            }
                                        }
                                    }
                                    numFilaCorreo = numFilaCorreo + 1;
                                }
                            }
                            else
                            {
                                
                            }
                        }
                    }
                    //int posMaxima = posFila.Length;
                }
                else
                {
                    Console.WriteLine("El tracker se encuentra vacío");
                }

                if(contadorEnCurso > 0)
                {
                    Console.WriteLine("====================================================================================");
                    Console.WriteLine("Esta ventana se cerrará automáticamente en cuatro (4) segundos");
                    Console.WriteLine("====================================================================================");
                    Console.WriteLine("© 2022 Saúl Candia");

                    //Contador para cierre de app consola:
                    Thread.Sleep(4000);
                    Environment.Exit(0);
                }
                else
                {
                    Console.WriteLine("Se encontraron " + atrasadoPosicionY.ToString() + " proyectos con estado ATRASADO.\n");
                    Console.WriteLine("Se encontraron " + porVencerPosicionY.ToString() + " proyectos que vencerán en los próximos "+ argumentoDias + " días.\n");
                    Console.WriteLine("Esta ventana se cerrará automáticamente en cuatro (4) segundos");
                    Console.WriteLine("====================================================================================");
                    Console.WriteLine("© 2022 Saúl Candia");
                    //Contador para cierre de app consola:
                    Thread.Sleep(4000);
                    Environment.Exit(0);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("===========================================================================================");
                Console.WriteLine("\nHUBO UN ERROR AL CONECTAR A LA API. Revise el mensaje a continuación para más detalles\n");
                Console.WriteLine("\n\n" + ex.Message.ToString()+ "\n\n");
            }
        }

        static Boolean SendMail(string nombreProyecto, DateTime fechaCompromiso, string nombreResponsable, string nombreClienteInterno, string correoResponsable, int estadoProyecto)
        {
            string estado1 = "";
            string estado2 = "";
            string mailBody = "";

            if (estadoProyecto == 1)
            {
                estado1 = "PRÓXIMO A VENCER";
                estado2 = "En curso (Vence pronto)";

                //GENERA EL CORREO Y LO ENVÍA
                mailBody = "<!DOCTYPE html><html><body><center><p><img src='cid:leoneraLogo'  width='100' height='75' /></p><h2>CORREO INFORMATIVO</h2><p>Hola, <b>" + nombreResponsable + "</b><br><br>Por medio de este correo, se le informa que, de acuerdo a nuestros registros uno de los proyectos del cual es responsable se encuentra <mark><b>" + estado1 + "</b></mark>." +
                        "<p><table>" +
                            "<tr>" +
                                "<td><b>Nombre del proyecto</b></td>" +
                                "<td style='color:green'>" + nombreProyecto + "</td>" +
                            "</tr>" +
                             "<tr>" +
                                "<td><b>Cliente Interno</b></td>" +
                                "<td style='color:green'>" + nombreClienteInterno + "</td>" +
                            "</tr>" +
                            "<tr>" +
                                "<td><b>Responsable</b></td>" +
                                "<td style='color:green'>" + nombreResponsable + "</td>" +
                            "</tr>" +
                            "<tr>" +
                                "<td><b>Fecha de compromiso</b></td>" +
                                "<td style='color:green'>" + fechaCompromiso.ToString("dd-MMM-yyyy") + "</td>" +
                            "</tr>" +
                            "<tr>" +
                                "<td><b>Estado</b></td>" +
                                "<td style='color:green'><b>" + estado2 + "</b></td>" +
                            "</tr>" +
                        "</table></p>" +
                        "<p><b>Revise sus fechas y verifique si puede llegar a la misma o solicitar una iteración (cambio).</b></p>" +
                        "<p><b>Por favor NO RESPONDA A ESTE CORREO</b> ya que fue generado de manera automática.</p>" +
                        "<p>Deseándole mucho éxito en sus proyectos, se despide atentamente, <b>GRUPO LEONERA.</b></p></center></body></html>";
            }
            else
            {
                estado1 = "ATRASADO";
                estado2 = "Vencido";

                mailBody = "<!DOCTYPE html><html><body><center><p><img src='cid:leoneraLogo'  width='100' height='75' /></p><h2>CORREO INFORMATIVO</h2><p>Hola, <b>" + nombreResponsable + "</b><br><br>Por medio de este correo, se le informa que, de acuerdo a nuestros registros uno de los proyectos del cual es responsable se encuentra <mark><b>" + estado1 + "</b></mark>." +
                           "<p><table>" +
                               "<tr>" +
                                   "<td><b>Nombre del proyecto</b></td>" +
                                   "<td>" + nombreProyecto + "</td>" +
                               "</tr>" +
                                "<tr>" +
                                   "<td><b>Cliente Interno</b></td>" +
                                   "<td>" + nombreClienteInterno + "</td>" +
                               "</tr>" +
                               "<tr>" +
                                   "<td><b>Responsable</b></td>" +
                                   "<td>" + nombreResponsable + "</td>" +
                               "</tr>" +
                               "<tr>" +
                                   "<td><b>Fecha de compromiso</b></td>" +
                                   "<td style='color:red'>" + fechaCompromiso.ToString("dd-MMM-yyyy") + "</td>" +
                               "</tr>" +
                               "<tr>" +
                                   "<td><b>Estado</b></td>" +
                                   "<td style='color:red'><b>" + estado2 + "</b></td>" +
                               "</tr>" +
                           "</table></p>" +
                           "<p style='color:red'><b>Se solicita regularizar la situación a la brevedad o en su defecto, conversar con solicitante para acordar nueva fecha de compromiso.</b></p>" +
                           "<p><b>Por favor NO RESPONDA A ESTE CORREO</b> ya que fue generado de manera automática.</p>" +
                           "<p>Deseándole mucho éxito en sus proyectos, se despide atentamente, <b>GRUPO LEONERA.</b></p></center></body></html> <br/>";
            }
            if (correoResponsable == "No hay")
            {
                return false;
            }
            else
            {
                
                try
                {
                    //GENERA EL CORREO Y LO ENVÍA
                   

                    MailMessage mm = new MailMessage();
                    mm.To.Add(correoResponsable);
                    //mm.CC.Add(search.emailRepresentante.ToString());
                    mm.From = new MailAddress("mensajero@leonera.cl", "Mensajes Leonera (NO RESPONDER)");
                    //mm.CC.Add(copiaGuardias);
                    mm.Subject = "Aviso de proyecto "+ estado1; //Asunto

                    //AGREGA IMÁGENES
                    AlternateView aw = AlternateView.CreateAlternateViewFromString(mailBody, null, MediaTypeNames.Text.Html);
                    //LinkedResource LOGO = new LinkedResource(Path.Combine(Directory.GetCurrentDirectory(), "Img/GrupoLeoneraLogo.jpg"), "image/jpg");
                    LinkedResource LOGO = new LinkedResource(Path.Combine(Directory.GetCurrentDirectory(), "D:/GrupoLeoneraLogo.jpg"), "image/jpg");
                    LOGO.ContentId = "leoneraLogo";
                    aw.LinkedResources.Add(LOGO);
                    mm.AlternateViews.Add(aw);
                    mm.Body = LOGO.ContentId;
                    SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                    smtp.EnableSsl = true;
                    NetworkCredential nc = new NetworkCredential("swdocumentalcg@leonera.cl", "L3oner4F0rest@l!");
                    smtp.Credentials = nc;
                    smtp.Send(mm);
                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: \n"+ex.Message.ToString());
                    return false;
                }
            }
            
          
        }

    }
}
