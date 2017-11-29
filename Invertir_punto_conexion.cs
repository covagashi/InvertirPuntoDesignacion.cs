// ConnectionPointDesignationReverse.cs
//
// Erweitert das Kontextmenü von 'Anschlussbezeichnungen',
// im Dialog 'Eigenschaften (Schaltzeichen): Allgemeines Betriebsmittel',
// um den Menüpunkt 'Reihenfolge drehen'.
// Es wird die Eingabe im Feld 'Anschlussbezeichnungen' automatisch gedreht.
//
// Copyright by Frank Schöneck, 2015
// letzte Änderung:
// V1.0.0, 04.03.2015, Frank Schöneck, Projektbeginn
//
// für Eplan Electric P8, ab V2.7
//
// Al hacer click derecho en designacion punto de conexion da la 
// posibilidad de invertir los numeros o letras.
//
// Traducido al español por Covagashi 2017
// Eplan Electric P8 V2.7

using System;
using System.Windows.Forms;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;

public class ConnectionPointDesignationReverse
{

	[DeclareMenu]
	public void ProjectCopyContextMenu()
	{
        //Elemento del menú contextual
		string menuText = getMenuText();
		Eplan.EplApi.Gui.ContextMenu oContextMenu = new Eplan.EplApi.Gui.ContextMenu();
		Eplan.EplApi.Gui.ContextMenuLocation oContextMenuLocation = new Eplan.EplApi.Gui.ContextMenuLocation("XDTDataDialog", "4006");
		oContextMenu.AddMenuItem(oContextMenuLocation, menuText, "ConnectionPointDesignationReverse", true, false);
	}

	[DeclareAction("ConnectionPointDesignationReverse")]
	public void Action()
	{
		try
		{
			string sSourceText = string.Empty;
			string sReturnText = string.Empty;
            string EplanCRLF = "¶";

            //Vaciar el portapapeles
			System.Windows.Forms.Clipboard.Clear();

			//copiar portapapeles
			CommandLineInterpreter oCLI = new CommandLineInterpreter();
			oCLI.Execute("GfDlgMgrActionIGfWind /function:SelectAll"); // seleccionar todo
			oCLI.Execute("GfDlgMgrActionIGfWind /function:Copy"); // copiar

			if (System.Windows.Forms.Clipboard.ContainsText())
			{
				sSourceText = System.Windows.Forms.Clipboard.GetText();
				if (sSourceText != string.Empty)
				{
					string[] sAnschlussbezeichnungen = sSourceText.Split(new string[] { EplanCRLF }, StringSplitOptions.None);

                    if (sAnschlussbezeichnungen.Length > 2) // Más de 2 puntos de conexión
					{
						Decider eDecision = new Decider();
						EnumDecisionReturn eAnswer = eDecision.Decide(EnumDecisionType.eYesNoDecision,
                            "Las designaciones de conexión están emparejadas?",
							"Invierte el orden",
							EnumDecisionReturn.eYES,
							EnumDecisionReturn.eYES,
							"ConnectionPointDesignationReverse",
							true,
							EnumDecisionIcon.eQUESTION);

						if (eAnswer == EnumDecisionReturn.eYES)
						{
                            // Reconstruir cadena de texto
							for (int i = 0; i < sAnschlussbezeichnungen.Length; i = i + 2)
							{
								sReturnText += sAnschlussbezeichnungen[i + 1] + EplanCRLF + sAnschlussbezeichnungen[i] + EplanCRLF;
							}
						}
						else
						{
                            //  Invierte el texto guardado en el array
							Array.Reverse(sAnschlussbezeichnungen);

                            // Reconstruir cadena de texto
							foreach (string sAnschluss in sAnschlussbezeichnungen)
							{
								sReturnText += sAnschluss + EplanCRLF;
							}
						}
					}
                    else // Solo 2 puntos de conexión
					{
						// Invierte el texto guardado en el array
						Array.Reverse(sAnschlussbezeichnungen);

                        // Reconstruir cadena de texto
						foreach (string sAnschluss in sAnschlussbezeichnungen)
						{
							sReturnText += sAnschluss + EplanCRLF;
						}
					}

                    // Devuelve el último carácter de nuevo
					sReturnText = sReturnText.Substring(0, sReturnText.Length - 1);

					//Insertar portapapeles
					System.Windows.Forms.Clipboard.SetText(sReturnText);
					oCLI.Execute("GfDlgMgrActionIGfWind /function:SelectAll"); // Seleccionar todo
					oCLI.Execute("GfDlgMgrActionIGfWindDelete"); // Borrar
					oCLI.Execute("GfDlgMgrActionIGfWind /function:Paste"); // Insertar
				}
			}
		}
		catch (System.Exception ex)
		{
			MessageBox.Show(ex.Message, "Invertir orden, Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
		return;
	}

	// Returns the menueitem text in the gui langueage if available.
	private string getMenuText()
	{
		MultiLangString muLangMenuText = new MultiLangString();
		muLangMenuText.SetAsString(
			"es_ES@INVERTIR ORDEN;" +
			"en_US@rotate order;"
			);

		ISOCode guiLanguage = new Languages().GuiLanguage;
		return muLangMenuText.GetString((ISOCode.Language)guiLanguage.GetNumber());
	}

}
