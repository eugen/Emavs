using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Emacs
{
    enum EmacsCommand
    {
        CancelCommand,

        SetMark,
        CharLeft,
        CharRight,
        WordLeft,
        WordRight,
        LineUp,
        LineDown,

        MoveBeginningOfLine,
        MoveEndOfLine,
        MoveBeginningOfFile,
        MoveEndOfFile,

        Yank,
        KillRegion,
        KillSave,
        KillLine,
    }
}
