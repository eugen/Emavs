﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;
using System.Reflection;

namespace Emacs
{
    class EmacsCommandHandler
    {
        public EnvDTE80.DTE2 App { get; set; }
        public EnvDTE.AddIn AddIn { get; set; }

        private string prevCommand;
        private bool emacsSelectionActive;
        private vsStartOfLineOptions prevStartOfLineOptions;
        private int prevScrollLineOptions;
        private Dictionary<string, Action> handlers;

        private TextSelection Selection {
            get {
                return (TextSelection)App.ActiveDocument.Selection;
            }
        }

        public List<string> SupportedCommands {
            get {
                return handlers.Keys.ToList();
            }
        }

        public EmacsCommandHandler() {
            handlers = GetType()
                .GetMethods(BindingFlags.NonPublic | BindingFlags.Instance)
                .Where(
                    m => m.Name.StartsWith("On") &&
                    m.ReturnType == typeof(void) &&
                    m.GetParameters().Length == 0)
                .ToDictionary(m => m.Name.Substring(2), m => (Action)Delegate.CreateDelegate(typeof(Action), this, m));
        }

        public void OnUnknownCommand() {
            prevCommand = null;
        }

        public void OnCommand(string cmdName) {
            Action action;
            if (handlers.TryGetValue(cmdName, out action)) {
                action();
            } else {
                System.Windows.Forms.MessageBox.Show("Unknown command: " + cmdName);
            }
            prevCommand = cmdName;
        }

        private void OnCancelCommand() {
            emacsSelectionActive = false;
            Selection.MoveToPoint(Selection.ActivePoint, false);
        }
        private void OnSetMark() {
            emacsSelectionActive = !emacsSelectionActive;
        }
        private void OnCharLeft() {
            Selection.CharLeft(emacsSelectionActive);
        }
        private void OnCharRight() {
            Selection.CharRight(emacsSelectionActive);
        }
        private void OnWordLeft() {
            Selection.WordLeft(emacsSelectionActive);
        }
        private void OnWordRight() {
            Selection.WordRight(emacsSelectionActive);
        }
        private void OnLineDown() {
            Selection.LineDown(emacsSelectionActive);
        }
        private void OnLineUp() {
            Selection.LineUp(emacsSelectionActive);
        }
        private void OnPageDown() {
            Selection.PageDown(emacsSelectionActive);
        }
        private void OnPageUp() {
            Selection.PageUp(emacsSelectionActive);
        }
        private void OnKillRegion() {
            if (emacsSelectionActive) {
                Selection.Cut();
                emacsSelectionActive = false;
            }
        }
        private void OnKillSave() {
            if (emacsSelectionActive) {
                Selection.Copy();
                Selection.MoveToPoint(Selection.ActivePoint, false);
                emacsSelectionActive = false;
            }
        }
        private void OnKillLine() {
            emacsSelectionActive = false;
            Selection.MoveToPoint(Selection.ActivePoint, false);
            Selection.EndOfLine(true);
            Selection.Cut();
        }
        private void OnYank() {
            Selection.Paste();
        }
        private void OnYankCycle() {
            App.ExecuteCommand("Edit.CycleClipboardRing");
            //TODO() { more emacs-ish kill-ring cycling
        }
        private void OnMoveBeginningOfFile() {
            Selection.StartOfDocument(emacsSelectionActive);
        }
        private void OnMoveEndOfFile() {
            Selection.EndOfDocument(emacsSelectionActive);
        }
        private void OnMoveBeginningOfLine() {
            if (prevCommand == "MoveBeginningOfLine") {
                if (prevStartOfLineOptions == vsStartOfLineOptions.vsStartOfLineOptionsFirstText) {
                    prevStartOfLineOptions = vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn;
                } else {
                    prevStartOfLineOptions = vsStartOfLineOptions.vsStartOfLineOptionsFirstText;
                }
            } else {
                prevStartOfLineOptions = vsStartOfLineOptions.vsStartOfLineOptionsFirstText;
            }
            var oldPos = Selection.ActivePoint.AbsoluteCharOffset;
            Selection.StartOfLine(prevStartOfLineOptions, emacsSelectionActive);

            // if we were already at the beginnig of the text, move to the beginning of the line
            if (oldPos == Selection.ActivePoint.AbsoluteCharOffset) {
                if (prevStartOfLineOptions == vsStartOfLineOptions.vsStartOfLineOptionsFirstText && !Selection.ActivePoint.AtStartOfLine) {
                    prevStartOfLineOptions = vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn;
                    Selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn, emacsSelectionActive);
                } else if (prevStartOfLineOptions == vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn) {
                    prevStartOfLineOptions = vsStartOfLineOptions.vsStartOfLineOptionsFirstText;
                    Selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText);
                }
            }
        }
        private void OnMoveEndOfLine() {
            Selection.EndOfLine(emacsSelectionActive);
        }
        private void OnRecenterTopBottom() {
            var value = 1;
            if (prevCommand == "RecenterTopBottom")
                value = (prevScrollLineOptions + 1) % 3;
            var commandRing = new []{"Edit.ScrollLineBottom","Edit.ScrollLineCenter","Edit.ScrollLineTop"};
            App.ExecuteCommand(commandRing[value]);
            prevScrollLineOptions = value;
        }
        private void OnForwardSexp() {
            throw new NotImplementedException();
        }
        private void OnBackwardSexp() {
            throw new NotImplementedException();
        }
    }
}