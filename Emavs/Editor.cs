using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using System.Linq;
using System.Collections.Generic;

namespace Emacs
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Editor : IDTExtensibility2, IDTCommandTarget
	{
        bool emacsSelectionActive = false;

		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Editor()
		{
		}

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			applicationObject = (DTE2)application;
			addInInstance = (AddIn)addInInst;

            Action<EmacsCommand> tryRegisterCommand = cmd => {
                var cmdName = cmd.ToString();
                try {
                    applicationObject.Commands.AddNamedCommand(addInInstance, cmdName, cmdName, cmdName, false);
                } catch {
                    // probably already registered; ignore
                }
            };
            ((EmacsCommand[])Enum.GetValues(typeof(EmacsCommand))).ToList().ForEach(tryRegisterCommand);
		}

           

		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		private DTE2 applicationObject;
		private AddIn addInInstance;
        private EmacsCommand? prevCommand;
        private vsStartOfLineOptions prevStartOfLineOptions;

        #region IDTCommandTarget Members

        public void Exec(string CmdName, vsCommandExecOption ExecuteOption, ref object VariantIn, ref object VariantOut, ref bool Handled) {

            EmacsCommand command;
            if(!CmdName.StartsWith("Emacs.Editor.") || 
               !Enum.TryParse<EmacsCommand>(CmdName.Substring("Emacs.Editor.".Length), out command)) {
                prevCommand = null; 
                return;
            }

            var sel = ((TextSelection)applicationObject.ActiveDocument.Selection);
            switch (command) {
                case EmacsCommand.CancelCommand:
                    emacsSelectionActive = false;
                    sel.MoveToPoint(sel.ActivePoint, false);
                    break;
                case EmacsCommand.SetMark:
                    emacsSelectionActive = !emacsSelectionActive;
                    break;
                case EmacsCommand.CharLeft:
                    sel.CharLeft(emacsSelectionActive);
                    break;
                case EmacsCommand.CharRight:
                    sel.CharRight(emacsSelectionActive);
                    break;
                case EmacsCommand.WordLeft:
                    sel.WordLeft(emacsSelectionActive);
                    break;
                case EmacsCommand.WordRight:
                    sel.WordRight(emacsSelectionActive);
                    break;
                case EmacsCommand.LineDown:
                    sel.LineDown(emacsSelectionActive);
                    break;
                case EmacsCommand.LineUp:
                    sel.LineUp(emacsSelectionActive);
                    break;
                case EmacsCommand.KillRegion:
                    if (emacsSelectionActive) {
                        sel.Cut();
                        emacsSelectionActive = false;
                    }
                    break;
                case EmacsCommand.KillSave:
                    if (emacsSelectionActive) {
                        sel.Copy();
                        sel.MoveToPoint(sel.ActivePoint, false);
                        emacsSelectionActive = false;
                    }
                    break;
                case EmacsCommand.KillLine:
                    emacsSelectionActive = false;
                    sel.MoveToPoint(sel.ActivePoint, false);
                    sel.EndOfLine(true);
                    sel.Cut();
                    break;
                case EmacsCommand.Yank:
                    applicationObject.ExecuteCommand("Edit.CycleClipboardRing");
                    //TODO: more emacs-ish kill-ring cycling
                    break;
                case EmacsCommand.MoveBeginningOfFile:
                    sel.StartOfDocument(emacsSelectionActive);
                    break;
                case EmacsCommand.MoveEndOfFile:
                    sel.EndOfDocument(emacsSelectionActive);
                    break;
                case EmacsCommand.MoveBeginningOfLine:
                    if (prevCommand == EmacsCommand.MoveBeginningOfLine) {
                        if (prevStartOfLineOptions == vsStartOfLineOptions.vsStartOfLineOptionsFirstText) {
                            prevStartOfLineOptions = vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn;
                        } else {
                            prevStartOfLineOptions = vsStartOfLineOptions.vsStartOfLineOptionsFirstText;
                        }
                    } else {
                        prevStartOfLineOptions = vsStartOfLineOptions.vsStartOfLineOptionsFirstText;
                    }
                    sel.StartOfLine(prevStartOfLineOptions, emacsSelectionActive);
                    break;
                case EmacsCommand.MoveEndOfLine:
                    sel.EndOfLine(emacsSelectionActive);
                    break;
                default:
                    prevCommand = null;
                    return;
            }
            prevCommand = command;
            Handled = true;
        }   

        public void QueryStatus(string CmdName, vsCommandStatusTextWanted NeededText, ref vsCommandStatus StatusOption, ref object CommandText) {
            StatusOption = vsCommandStatus.vsCommandStatusEnabled | vsCommandStatus.vsCommandStatusSupported;
        }

        #endregion
    }
}