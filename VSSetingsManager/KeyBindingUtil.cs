using EnvDTE;
using System;

namespace VSSettingsManager
{
    class KeyBindingUtil
    {
        private static void AddKeyBinding(Commands cmds, string vsCommandName, string scope, string keyBinding)
        {
            Command command = cmds.Item(vsCommandName);
            command.Bindings = (object)AppendKeyboardBinding(command, $"{scope}::{keyBinding}");
        }

        private static void ReplaceKeyBinding(Commands cmds, string vsCommandName, string scope, string keyBinding)
        {
            Command command = cmds.Item(vsCommandName);
            // Build new array with just the new binding.
            command.Bindings = new object[] { $"{scope}::{keyBinding}" };
        }

        private static object[] AppendKeyboardBinding(Command command, string keyboardBindingDefn)
        {
            object[] oldBindings = (object[])command.Bindings;

            // Check that keyboard binding is not already there
            for (int i = 0; i < oldBindings.Length; i++)
            {
                if (keyboardBindingDefn.Equals(oldBindings[i]))
                {
                    // Exit early and return the existing bindings array if new keyboard binding is already there
                    return oldBindings;
                }
            }

            // Build new array with all the old bindings, plus the new one.
            object[] newBindings = new object[oldBindings.Length + 1];
            Array.Copy(oldBindings, newBindings, oldBindings.Length);
            newBindings[newBindings.Length - 1] = keyboardBindingDefn;
            return newBindings;
        }

    }
}
