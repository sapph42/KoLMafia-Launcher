using System;
using System.Linq;
using System.Windows.Forms;

static void Main(string[] args) {
    if (args.Length == 0) {
        bool noLaunch = false;
        bool killOnUpdate = false;
        bool silent = false;
    } else {
        if (args.Contains("--noLaunch",StringComparer.CurrentCultureIgnoreCase)) {
            bool noLaunch = true;
        } else {
            bool noLaunch = false;
        }
        if (args.Contains("--killOnUpdate",StringComparer.CurrentCultureIgnoreCase)) {
            bool killOnUpdate = true;
        } else {
            bool killOnUpdate = false;
        }
        if (args.Contains("--silent",StringComparer.CurrentCultureIgnoreCase)) {
            bool silent = true;
        } else {
            bool silent = false;
        }                  
    }
}