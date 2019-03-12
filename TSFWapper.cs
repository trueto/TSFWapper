using System;  
using System.Text;  
using TSF.TypeLib;  
using System.Runtime.InteropServices;  
	
namespace TestIME  
{  
    class TSFWapper  
    {  
        // Gets the name of the current IME  
        public string GetName()  
        {  
            var tITfInputProcessorProfiles = Type.GetTypeFromCLSID(new Guid(TSF.TypeLib.CLSIDs.CLSID_TF_InputProcessorProfiles));  
            var obj = Activator.CreateInstance(tITfInputProcessorProfiles);  
            try  
            {  
                var inputProcessorProfiles = obj as TSF.TypeLib.ITfInputProcessorProfiles;  
                var inputProcessorProfileManager = (TSF.TypeLib.ITfInputProcessorProfileMgr)inputProcessorProfiles;  
                TSF.TypeLib.TF_INPUTPROCESSORPROFILE prof;  
                string description = "";  

                Guid GUID_TFCAT_TIP_KEYBOARD = TSF.TypeLib.Guids.GUID_TFCAT_TIP_KEYBOARD;  
                inputProcessorProfileManager.GetActiveProfile(ref GUID_TFCAT_TIP_KEYBOARD, out prof);  
                if (prof.dwProfileType == TSF.TypeLib.TF_PROFILETYPE.TF_PROFILETYPE_INPUTPROCESSOR)  
                {  
                    inputProcessorProfiles.GetLanguageProfileDescription(ref prof.clsid, prof.langid, ref prof.guidProfile, out description);  
                }  
                else if (prof.dwProfileType == TF_PROFILETYPE.TF_PROFILETYPE_KEYBOARDLAYOUT)  
                {  
                    description = "KEYBOARD";  
                }  
                return description;  

            }  
            finally   
            {  

                if (obj != null) Marshal.ReleaseComObject(obj);  
            }  
        }  

        // Switch to the specified IME  
        public void SwitchIME(string keyword)  
        {  
            if (keyword==GetName())  
            {  
                return;  
            }  
            var HKL = GetKeyboardLayout(0);  
            Type.GetTypeFromCLSID(new Guid(TSF.TypeLib.CLSIDs.CLSID_TF_ThreadMgr));  
            var tITfInputProcessorProfiles = Type.GetTypeFromCLSID(new Guid(TSF.TypeLib.CLSIDs.CLSID_TF_InputProcessorProfiles));  
            var obj = Activator.CreateInstance(tITfInputProcessorProfiles);  
            try  
            {  
                var inputProcessorProfiles = obj as TSF.TypeLib.ITfInputProcessorProfiles;  
                var inputProcessorProfileManager = (TSF.TypeLib.ITfInputProcessorProfileMgr)inputProcessorProfiles;  
                TSF.TypeLib.TF_INPUTPROCESSORPROFILE prof;  
                string description = "";  
                uint fetch;  
                TSF.TypeLib.IEnumTfInputProcessorProfiles pp;  
                TSF.InteropTypes.LangID lid = new TSF.InteropTypes.LangID((ushort)HKL);  
                inputProcessorProfileManager = (TSF.TypeLib.ITfInputProcessorProfileMgr)inputProcessorProfiles;  
                if (inputProcessorProfileManager.EnumProfiles(lid, out pp).Code == 0 && pp != null)  
                {  
                    while (pp.Next(1, out prof, out fetch).Code == 0)  
                    {  
                        if (prof.dwProfileType == TSF.TypeLib.TF_PROFILETYPE.TF_PROFILETYPE_INPUTPROCESSOR)  
                        {  
                            inputProcessorProfiles.GetLanguageProfileDescription(ref prof.clsid, prof.langid, ref prof.guidProfile, out description);  
                            if (description==keyword)  
                            {   
                                inputProcessorProfiles.ActivateLanguageProfile(ref prof.clsid, prof.langid, prof.guidProfile);  
                            };  
                        }  
                    }  
                }  

            }  
            finally  
            {  
                if (obj != null) Marshal.ReleaseComObject(obj);  
            }  
        }  
    }  
}
