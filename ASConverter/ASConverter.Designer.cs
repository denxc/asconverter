﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан программой.
//     Исполняемая версия:4.0.30319.42000
//
//     Изменения в этом файле могут привести к неправильной работе и будут потеряны в случае
//     повторной генерации кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ASConverter {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "14.0.0.0")]
    internal sealed partial class ASConverter : global::System.Configuration.ApplicationSettingsBase {
        
        private static ASConverter defaultInstance = ((ASConverter)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new ASConverter())));
        
        public static ASConverter Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        public global::System.Collections.Specialized.StringCollection SourceFiles {
            get {
                return ((global::System.Collections.Specialized.StringCollection)(this["SourceFiles"]));
            }
            set {
                this["SourceFiles"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        public global::System.Collections.Specialized.StringCollection DestFiles {
            get {
                return ((global::System.Collections.Specialized.StringCollection)(this["DestFiles"]));
            }
            set {
                this["DestFiles"] = value;
            }
        }
    }
}
