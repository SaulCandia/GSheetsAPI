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
using System.Text;

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
            int columnas = 8;

            if (args != null || args.Length == 0)
            {
                try
                {
                    argumentoDias = int.Parse(args[0].ToString());
                }
                catch (Exception)
                {
                    Console.WriteLine("Parámetro debe ser un número entero (Días de anticipación)");
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
            Console.WriteLine("Saúl Candia - Área Transformación Digital, Grupo Leonera.");
            //String spreadsheetId = "1ekSXRUep9T0U7ZuVCDz0fJ7X_trGZp8JmrzL7b8OJb4"; /*ID copia Tracker*/
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

                //Declaración de variables
                int numFila = 1;
                int limite = values.Count;
                var hoy = DateTime.Today;
                int contadorElementos = 0;
                Boolean result = false;

                //Declaración de matriz bidimensional [cantidad elementos encontrados en Hoja, 7]
                //      0                   1                   2                       3                   4                   5                   6               7
                //[nombreProyecto], [fechaCompromiso], [nombreClienteInterno], [nombreResponsable], [correoResponsable], [estadoProyecto], [rowPositionIndex], [diasDiferencia]
                string[,] datosTracker = new string[limite, columnas];

                //PARTE 1: REVISA EL ESTADO DE TAREAS Y DETERMINA POSICIÓN DE FILA X EN LAS MISMAS

                if (values != null && values.Count > 0)
                {
                    int rowPosition = 0;
                    foreach (var rowEstado in values)
                    {
                        datosTracker[rowPosition, 5] = rowEstado[0].ToString();
                        datosTracker[rowPosition, 6] = numFila.ToString();
                        rowPosition = rowPosition + 1;
                        numFila = numFila + 1;
                    }
                }
                else
                {
                    Console.WriteLine("El tracker se encuentra vacío");
                }

                //PARTE 2: REVISA EL NOMBRE DE LAS TAREAS, LA FECHA DE COMPROMISO INICIAL Y LOS DIAS DE DIFERENCIA ENTRE HOY Y ESA FECHA, SEGÚN POSICIÓN DE FILA

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nBuscando nombre de proyectos, espere por favor... \n");


                range = "Tracker!G:H";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                response = request.Execute();
                values = response.Values;

                if (values != null && values.Count > 0)
                {
                    int rowPosition = 0;
                    DateTime fechaCompromiso = DateTime.Today;
                    foreach (var rowNombreTareas in values)
                    {
                        try
                        {
                            datosTracker[rowPosition, 0] = rowNombreTareas[0].ToString();
                            datosTracker[rowPosition, 1] = rowNombreTareas[1].ToString();

                            if (rowNombreTareas[1].ToString() != "Fecha Comp Inicial")
                            {
                                //Cálculo de fechas
                                fechaCompromiso = Convert.ToDateTime(rowNombreTareas[1].ToString());
                                TimeSpan resta = fechaCompromiso.Date - hoy;
                                int difDias = int.Parse(resta.Days.ToString());
                                datosTracker[rowPosition, 7] = difDias.ToString();
                            }
                            else
                            {
                                datosTracker[rowPosition, 7] = "0";
                            }
                        }
                        catch (Exception)
                        {
                            datosTracker[rowPosition, 1] = "";
                        }

                        rowPosition = rowPosition + 1;
                    }
                }
                else
                {
                    Console.WriteLine("El tracker se encuentra vacío");
                }


                //PARTE 3: REVISA AL CLIENTE INTERNO Y AL RESPONSABLE

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nRevisando cliente interno y responsable de la tarea, espere por favor... \n");


                range = "Tracker!K:L";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                response = request.Execute();
                values = response.Values;

                if (values != null && values.Count > 0)
                {
                    int rowPosition = 0;
                    foreach (var rowClienteInterno in values)
                    {
                        datosTracker[rowPosition, 2] = rowClienteInterno[0].ToString();
                        datosTracker[rowPosition, 3] = rowClienteInterno[1].ToString();
                        rowPosition = rowPosition + 1;
                    }
                }
                else
                {
                    Console.WriteLine("El tracker se encuentra vacío");
                }


                //PARTE 4: REVISA CORREOS DE LOS RESPONSABLES

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nRevisando listado de correos, espere por favor... \n");


                range = "DataCorreo!A:B";
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);
                response = request.Execute();
                values = response.Values;

                if (values != null && values.Count > 0)
                {
                    int w = 0;
                    foreach (var rowCorreoResponsable in values)
                    {
                        for (w = 1; w < limite; w++)
                        {
                            if (rowCorreoResponsable[0].ToString() == datosTracker[w, 3].ToString())
                            {
                                datosTracker[w, 4] = rowCorreoResponsable[1].ToString();
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("El tracker se encuentra vacío");
                }

                //PARTE 5: SELECCIÓN DE TAREAS VENCIDAS Y POR VENCER SEGÚN CANTIDAD DE DIAS

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nEvaluando tareas vencidas y por vencer, espere por favor... \n");

                int i;
                int diferenciaDias;
                for (i = 1; i < limite; i++)
                {
                    diferenciaDias = 0;
                    if (datosTracker[i, 5].ToString() == "En Curso" || datosTracker[i, 5].ToString() == "Atrasado")
                    {
                        diferenciaDias = int.Parse(datosTracker[i, 7].ToString());
                        if (diferenciaDias <= argumentoDias)
                        {
                            contadorElementos = contadorElementos + 1;
                        }
                    }
                }


                //Declaración de matriz bidimensional con datos en limpio [cantidad elementos encontrados en Hoja, 7]
                //      0                   1                   2                       3                   4                   5                   6               7
                //[nombreProyecto], [fechaCompromiso], [nombreClienteInterno], [nombreResponsable], [correoResponsable], [estadoProyecto], [rowPositionIndex], [diasDiferencia]
                string[,] datosTrackerFinal = new string[contadorElementos, columnas];


                int a = 0;
                int b = 0;
                int ultimaposicion = 0;

                for (i = 1; i < limite; i++)
                {
                    diferenciaDias = int.Parse(datosTracker[i, 7].ToString());
                    if (diferenciaDias <= argumentoDias)
                    {
                        if (datosTracker[i, 5].ToString() == "En Curso" || datosTracker[i, 5].ToString() == "Atrasado")
                        {
                            for (a = ultimaposicion; a < contadorElementos; a++)
                            {
                                for (b = 0; b < columnas; b++)
                                {
                                    datosTrackerFinal[a, b] = datosTracker[i, b].ToString();
                                }
                                ultimaposicion++; ;
                                break;
                            }
                        }
                    }
                }

                //PARTE 6: Enviar el correo:

                Console.WriteLine("====================================================================================");
                Console.WriteLine("\nEnviando correos a responsables, espere por favor \n");

                //Envia correo

                result = SendMail(datosTrackerFinal, contadorElementos);

                if (result == true)
                {
                    Console.WriteLine("\nEl correo fue enviado exitosamente. :-)");
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
                    Console.WriteLine("\nEl correo no fue enviado.\nProbablemente se deba a que no hay correo definido para el responsable o porque el mismo no existe.");
                    //Contador para cierre de app consola:
                    Thread.Sleep(10000);
                    Environment.Exit(0);
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("===========================================================================================");
                Console.WriteLine("\nHUBO UN ERROR AL CONECTAR A LA API. Revise el mensaje a continuación para más detalles\n");
                Console.WriteLine("\n\n" + ex.Message.ToString() + "\n\n");
            }


            



                //if(contadorEnCurso > 0)
                //{
                //    Console.WriteLine("====================================================================================");
                //    Console.WriteLine("Esta ventana se cerrará automáticamente en cuatro (4) segundos");
                //    Console.WriteLine("====================================================================================");
                //    Console.WriteLine("© 2022 Saúl Candia");

                //    //Contador para cierre de app consola:
                //    Thread.Sleep(4000);
                //    Environment.Exit(0);
                //}
                //else
                //{
                //    Console.WriteLine("Se encontraron " + atrasadoPosicionY.ToString() + " proyectos con estado ATRASADO.\n");
                //    Console.WriteLine("Se encontraron " + porVencerPosicionY.ToString() + " proyectos que vencerán en los próximos "+ argumentoDias + " días.\n");
                //    Console.WriteLine("Esta ventana se cerrará automáticamente en cuatro (4) segundos");
                //    Console.WriteLine("====================================================================================");
                //    Console.WriteLine("© 2022 Saúl Candia");
                //    //Contador para cierre de app consola:
                //    Thread.Sleep(4000);
                //    Environment.Exit(0);
          
        }

        //static Boolean SendMail(string nombreProyecto, DateTime fechaCompromiso, string nombreResponsable, string nombreClienteInterno, string correoResponsable, int estadoProyecto)
        static Boolean SendMail(string[,] listaDefinitiva, int contadorElementos)
        {
            int a = 0;
            int b = 0;
            int i = 0;
            string nombreResponsable = "";
            string correoResponsable;

            string[] nombreProyecto = new string[contadorElementos];
            string[] fechaCompromiso = new string[contadorElementos];
            string[] nombreClienteInterno = new string[contadorElementos];
            string[] estadoTarea = new string[contadorElementos];
            string[] diferenciaDias = new string[contadorElementos];

            for (a = 0; a < contadorElementos; a++)
            {
                if (nombreResponsable != listaDefinitiva[a, 3].ToString() || listaDefinitiva[a, 3].ToString() != "ENVIADO" || nombreResponsable != "ENVIADO")
                {
                    nombreResponsable = listaDefinitiva[a, 3].ToString();
                    correoResponsable = listaDefinitiva[a, 4].ToString();

                    for (b=0; b < contadorElementos; b++)
                    {
                        if (listaDefinitiva[b, 3].ToString() == nombreResponsable)
                        {
                            nombreProyecto[b] = listaDefinitiva[b, 0].ToString();
                            fechaCompromiso[b] = listaDefinitiva[b, 1].ToString();
                            nombreClienteInterno[b] = listaDefinitiva[b, 2].ToString();
                            estadoTarea[b] = listaDefinitiva[b, 5].ToString();
                            diferenciaDias[b] = listaDefinitiva[b, 7].ToString();
                        } 
                    }

                    // Aqui se genera y envia el correo acumulado:
                    try
                    {
                        StringBuilder builder = new StringBuilder();

                        builder.Append("<!DOCTYPE html><html><body><center><p><img src='cid:leoneraLogo'  width='100' height='75' /></p><h2>CORREO INFORMATIVO</h2><p>Hola, <b>" + nombreResponsable + "</b><br><br>Por medio de este correo, se le informa que, de acuerdo a nuestros registros usted tiene proyectos que deben ser revisados porque están por vencer o ya vencieron.");

                        for (i = 0; i < contadorElementos; i++)
                        {
                            if(nombreProyecto[i] != null)
                            {
                                builder.Append("<hr style='width: 600px; color:gray;'>" +
                                    "<p><table style='table-layout: fixed;'><tr>" +
                                          "<td style='width: 200px;'><b>Nombre del proyecto</b></td>" +
                                          "<td style='width: 400px;'>" + nombreProyecto[i].ToString() + "</td>" +
                                      "</tr>" +
                                       "<tr>" +
                                          "<td style='width: 200px;'><b>Cliente Interno</b></td>" +
                                          "<td style='width: 400px;'>" + nombreClienteInterno[i].ToString() + "</td>" +
                                      "</tr>" +
                                      "<tr>" +
                                          "<td style='width: 200px;'><b>Responsable</b></td>" +
                                          "<td style='width: 400px;'>" + nombreResponsable + "</td>" +
                                      "</tr>" +
                                      "<tr>" +
                                          "<td style='width: 200px;'><b>Fecha de compromiso</b></td>" +
                                          "<td style='width: 400px; color:blue;'>" + fechaCompromiso[i].ToString() + "</td>" +
                                      "</tr>" +
                                      "<tr>" +
                                          "<td style='width: 200px;'><b>Estado</b></td>" +
                                          "<td style='width: 400px; color:red;'><b>" + estadoTarea[i] + "</b></td>" +
                                      "</tr>" +
                                  "</table></p>");
                            }
                           
                        }
                        builder.Append("<hr style='width: 600px; color:gray;'>"); 
                        builder.Append("<p style='color:green'><b>Se solicita regularizar la situación a la brevedad o en su defecto, conversar con solicitante para acordar nueva fecha de compromiso.</b></p>");
                        builder.Append("<p><b>Por favor NO RESPONDA A ESTE CORREO</b> ya que fue generado de manera automática.</p>");
                        builder.Append("<p>Deseándole mucho éxito en sus proyectos, se despide atentamente, <b>GRUPO LEONERA.</b></p></center></body></html> <br/>");
                   
                        //GENERA EL CORREO Y LO ENVÍA


                        MailMessage mm = new MailMessage();
                        mm.To.Add(correoResponsable);
                        //mm.CC.Add(search.emailRepresentante.ToString());
                        mm.From = new MailAddress("mensajero@leonera.cl", "Mensajes Leonera (NO RESPONDER)");
                        //mm.CC.Add(copiaGuardias);
                        mm.Subject = "Alerta de Tracker";
                        mm.Body = builder.ToString();

                        //AGREGA IMÁGENES
                        AlternateView aw = AlternateView.CreateAlternateViewFromString(mm.Body, null, MediaTypeNames.Text.Html);
                        LinkedResource LOGO = new LinkedResource(Path.Combine(Directory.GetCurrentDirectory(), "/GrupoLeoneraLogo.jpg"), "image/jpg");
                        //LinkedResource LOGO = new LinkedResource(Path.Combine(Directory.GetCurrentDirectory(), "D:/GrupoLeoneraLogo.jpg"), "image/jpg");
                        LOGO.ContentId = "leoneraLogo";
                        aw.LinkedResources.Add(LOGO);
                        mm.AlternateViews.Add(aw);
                        mm.Body = LOGO.ContentId;
                        SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                        smtp.EnableSsl = true;
                        NetworkCredential nc = new NetworkCredential("swdocumentalcg@leonera.cl", "L3oner4F0rest@l!");
                        smtp.Credentials = nc;
                        smtp.Send(mm);

                        //BORRA DATOS YA ENVIADOS


                        for (b = 0; b < contadorElementos; b++)
                        {
                            if (listaDefinitiva[b, 3].ToString() == nombreResponsable)
                            {
                                nombreProyecto[b] = null;
                                fechaCompromiso[b] = null;
                                nombreClienteInterno[b] = null;
                                estadoTarea[b] = null;
                                diferenciaDias[b] = null;
                                listaDefinitiva[b, 3] = "ENVIADO";
                            }
                        }
                        nombreResponsable = "ENVIADO";
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: \n" + ex.Message.ToString());
                        return false;
                    }

                }

                else
                {
                    //nombreProyecto[a] = listaDefinitiva[a, 0].ToString();
                    //fechaCompromiso[a] = listaDefinitiva[a, 1].ToString();
                    //nombreClienteInterno[a] = listaDefinitiva[a, 2].ToString();
                    //estadoTarea[a] = listaDefinitiva[a, 5].ToString();
                    //diferenciaDias[a] = listaDefinitiva[a, 7].ToString();
                    a = a + 0;
                    //nombreResponsable = "";
                }
            }
            return true;
        }

    }
}